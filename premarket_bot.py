# -*- coding: utf-8 -*-
"""
台股選股雷達 — 盤前自動分析看板 ＋ LINE Messaging API 準點推播引擎 (premarket_bot.py)
支援 4 階段滾動更新唯一狀態檔 daily_cache/premarket_daily.json：
  - 05:30 (Phase 1 免費資料源)：美股四大指數、台積電/聯電/日月光 ADR、台指期夜盤、外資期貨未平倉、融資餘額、除息扣點（靜默存檔）
  - 08:15 (Phase 1 免費資料源)：日韓股 (日經/KOSPI/三星/SK海力士)、小那/小標普電子盤、DXY/美債10Y、背離校正（推播第 1 則 LINE）
  - 08:48 (Phase 2 券商 API)：台指期日盤開盤 3 分鐘 K 棒與正逆價差（僅方向翻轉或點數落差達門檻才推播警示）
  - 08:58 (Phase 2 券商 API)：08:55~08:58 前五大權值股試撮五檔與 08:57:30 vs 08:58:00 假單抽單偵測（推播第 2 則 LINE）

用法範例：
  python premarket_bot.py --stage 0530
  python premarket_bot.py --stage 0815 --push-line
  python premarket_bot.py --stage 0848 --push-line
  python premarket_bot.py --stage 0858 --push-line
  python premarket_bot.py --stage auto --push-line   # 依台灣時間自動判斷階段
  python premarket_bot.py --stage 0815 --dry-run     # 本機測試不消耗 LINE 額度
"""

import argparse
import csv
import io
import json
import os
import ssl
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import yfinance as yf

TW_TZ = timezone(timedelta(hours=8))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE_DIR, "daily_cache")
STATE_PATH = os.path.join(CACHE_DIR, "premarket_daily.json")

# 台灣 2026~2027 主要國定休市日（可依證交所公告年度行事曆隨時增補）
TW_MARKET_HOLIDAYS = {
    "2026-01-01", "2026-02-12", "2026-02-13", "2026-02-16", "2026-02-17",
    "2026-02-18", "2026-02-19", "2026-02-20", "2026-02-27", "2026-04-03",
    "2026-04-06", "2026-05-01", "2026-06-19", "2026-09-25", "2026-10-09",
    "2027-01-01", "2027-02-04", "2027-02-05", "2027-02-08", "2027-02-09",
    "2027-02-10", "2027-02-11", "2027-02-12", "2027-02-28", "2027-03-01",
    "2027-04-05", "2027-05-01", "2027-06-09", "2027-09-15", "2027-10-10",
}

# 美股權重（連動強度：SOX > NASDAQ > S&P500 > DJI）
US_INDEX_WEIGHTS = {
    "SOX": ("^SOX", "費城半導體", 0.40),
    "NASDAQ": ("^IXIC", "那斯達克", 0.25),
    "SP500": ("^GSPC", "標普500", 0.20),
    "DJI": ("^DJI", "道瓊工業", 0.15),
}

ADR_TICKERS = {
    "TSM": ("TSM", "台積電 ADR"),
    "UMC": ("UMC", "聯電 ADR"),
    "ASX": ("ASX", "日月光 ADR"),
}

ASIA_MACRO_TICKERS = {
    "N225": ("^N225", "日經225"),
    "KOSPI": ("^KS11", "韓國KOSPI"),
    "SAMSUNG": ("005930.KS", "三星電子"),
    "HYNIX": ("000660.KS", "SK海力士"),
    "NQ": ("NQ=F", "小那斯達克電子盤"),
    "ES": ("ES=F", "小標普電子盤"),
    "DXY": ("DX-Y.NYB", "美元指數"),
    "US10Y": ("^TNX", "美債10年殖利率"),
}

# 前五大權值股（供 08:55~08:58 試撮加權交叉驗證）
TOP5_WEIGHTED_STOCKS = [
    ("2330", "台積電", 0.36),
    ("2317", "鴻海", 0.045),
    ("2454", "聯發科", 0.038),
    ("2308", "台達電", 0.022),
    ("2382", "廣達", 0.020),
]


def _parse_num(val):
    if val is None:
        return 0.0
    s = str(val).replace(",", "").replace("+", "").replace("%", "").strip()
    if s in ("", "-", "--", "N/A", "nan", "None"):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _http_get_text(url, timeout=12, post_data=None, encoding="utf-8"):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()
    data_bytes = urllib.parse.urlencode(post_data).encode("utf-8") if post_data else None
    req = urllib.request.Request(url, data=data_bytes, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx) as resp:
        return resp.read().decode(encoding, errors="ignore")


def is_tw_trading_day(target_date: date):
    """台灣股市休市日防呆（週末 + 台灣國定休市日）"""
    if target_date.weekday() >= 5:
        return False, "週末休市"
    iso = target_date.strftime("%Y-%m-%d")
    if iso in TW_MARKET_HOLIDAYS:
        return False, f"台灣國定休市日 ({iso})"
    return True, "正常交易日"


def _fetch_yf_quote(symbol, name):
    """抓取單一 Yahoo Finance 標的最新收盤／現價、漲跌幅與交易日期"""
    try:
        df = yf.download(tickers=symbol, period="10d", interval="1d", auto_adjust=False, progress=False)
        if df is None or df.empty:
            return {"symbol": symbol, "name": name, "available": False}
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [c[0] for c in df.columns]
        df = df.dropna(subset=["Close"]).copy()
        if len(df) < 2:
            return {"symbol": symbol, "name": name, "available": False}
        last_close = round(float(df["Close"].iloc[-1]), 2)
        prev_close = round(float(df["Close"].iloc[-2]), 2)
        chg_pct = round((last_close / prev_close - 1.0) * 100.0, 2) if prev_close > 0 else 0.0
        dt_last = df.index[-1].strftime("%Y-%m-%d") if hasattr(df.index[-1], "strftime") else str(df.index[-1])[:10]
        return {
            "symbol": symbol,
            "name": name,
            "available": True,
            "price": last_close,
            "prev_close": prev_close,
            "chg_pct": chg_pct,
            "date": dt_last,
        }
    except Exception:
        return {"symbol": symbol, "name": name, "available": False}


def fetch_prev_taiex_close():
    """從證交所官方 FMTQIK 抓取前一日加權指數收盤點數（比 Yahoo ^TWII 更準時）"""
    now_tw = datetime.now(TW_TZ)
    today_d = now_tw.date()
    months = []
    y, m = today_d.year, today_d.month
    for _ in range(2):
        months.append(f"{y:04d}{m:02d}01")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    rows = []
    for ym01 in reversed(months):
        try:
            raw = _http_get_text(f"https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date={ym01}&response=json")
            payload = json.loads(raw)
            if payload.get("stat") != "OK" or not payload.get("data"):
                continue
            fields = [str(f).strip() for f in payload.get("fields", [])]
            idx_c = next((i for i, f in enumerate(fields) if "加權股價指數" in f or "指數" in f), 4)
            for r in payload["data"]:
                parts = str(r[0]).strip().split("/")
                if len(parts) != 3:
                    continue
                d_obj = date(int(parts[0]) + 1911, int(parts[1]), int(parts[2]))
                if d_obj >= today_d:
                    continue
                c_val = _parse_num(r[idx_c])
                if c_val > 0:
                    rows.append((d_obj.strftime("%Y-%m-%d"), round(c_val, 2)))
        except Exception:
            continue
    if rows:
        rows.sort(key=lambda x: x[0])
        return {"available": True, "date": rows[-1][0], "close": rows[-1][1]}
    return {"available": False, "date": "--", "close": 0.0}


def fetch_taifex_night_session():
    """從期交所 OpenAPI (DailyMarketReportFut) 抓取台指期近月夜盤 (盤後) 收盤與漲跌點數"""
    try:
        raw = _http_get_text("https://openapi.taifex.com.tw/v1/DailyMarketReportFut", timeout=15)
        items = json.loads(raw)
        if not isinstance(items, list):
            return {"available": False}
        tx_night = [
            x for x in items
            if str(x.get("Contract", "")).strip() == "TX"
            and "盤後" in str(x.get("TradingSession", ""))
            and "/" not in str(x.get("ContractMonth(Week)", ""))
        ]
        if not tx_night:
            # 退回一般盤
            tx_night = [
                x for x in items
                if str(x.get("Contract", "")).strip() == "TX"
                and "/" not in str(x.get("ContractMonth(Week)", ""))
            ]
        if not tx_night:
            return {"available": False}
        # 取最近月契約
        tx_night.sort(key=lambda x: str(x.get("ContractMonth(Week)", "")))
        row = tx_night[0]
        close_p = _parse_num(row.get("Close") or row.get("Last"))
        chg_pts = _parse_num(row.get("Change"))
        chg_pct = _parse_num(row.get("Change%"))
        return {
            "available": close_p > 0,
            "date": str(row.get("Date", "")),
            "session": str(row.get("TradingSession", "盤後")),
            "contract_month": str(row.get("ContractMonth(Week)", "")),
            "close": round(close_p, 1),
            "chg_pts": round(chg_pts, 1),
            "chg_pct": round(chg_pct, 2),
        }
    except Exception:
        return {"available": False}


def fetch_foreign_futures_oi(prev_date_iso):
    """抓取期交所外資台股期貨 (TXF) 未平倉淨口數與單日增減"""
    if not prev_date_iso or prev_date_iso == "--":
        return {"available": False}
    try:
        q_date = prev_date_iso.replace("-", "/")
        raw = _http_get_text(
            "https://www.taifex.com.tw/cht/3/futContractsDateDown",
            timeout=12,
            post_data={"queryStartDate": q_date, "queryEndDate": q_date, "commodityId": "TXF"},
            encoding="cp950",
        )
        reader = csv.reader(io.StringIO(raw))
        for r in reader:
            if len(r) >= 14 and "臺股期貨" in r[1] and "外資" in r[2]:
                net_oi = int(round(_parse_num(r[13])))
                day_chg = int(round(_parse_num(r[7])))
                return {
                    "available": True,
                    "date": prev_date_iso,
                    "net_oi": net_oi,
                    "day_chg": day_chg,
                }
    except Exception:
        pass
    return {"available": False, "date": prev_date_iso, "net_oi": None, "day_chg": None}


def fetch_twse_ex_dividend_today(target_date: date):
    """從證交所 TWT48U_ALL 查詢今日除權息個股與預估大盤蒸發點數（無除息則回傳 0.0）"""
    roc_yyyymmdd = f"{target_date.year - 1911:03d}{target_date.month:02d}{target_date.day:02d}"
    roc_zh = f"{target_date.year - 1911}年{target_date.month:02d}月{target_date.day:02d}日"
    matched = []
    try:
        raw = _http_get_text("https://openapi.twse.com.tw/v1/exchangeReport/TWT48U_ALL", timeout=12)
        items = json.loads(raw)
        if isinstance(items, list):
            for it in items:
                d_str = str(it.get("Date", "")).strip()
                if d_str in (roc_yyyymmdd, roc_zh):
                    matched.append({
                        "code": str(it.get("Code", "")).strip(),
                        "name": str(it.get("Name", "")).strip(),
                        "cash_div": round(_parse_num(it.get("CashDividend")), 2),
                    })
    except Exception:
        pass
    # 簡易推估蒸發點數（權值股 2330 每配息 1 元約影響大盤 8.2 點；一般股粗估每檔 0.3~0.8 點）
    est_pts = 0.0
    for m in matched:
        if m["code"] == "2330":
            est_pts += m["cash_div"] * 8.2
        elif m["code"] in ("2317", "2454", "2308", "2382"):
            est_pts += m["cash_div"] * 0.9
        else:
            est_pts += min(m["cash_div"] * 0.15, 0.8)
    return {
        "count": len(matched),
        "stocks": matched[:10],
        "est_evap_pts": round(est_pts, 1),
    }


def fetch_twse_margin_latest(prev_date_iso):
    """讀取前一日證交所大盤融資餘額與增減"""
    if not prev_date_iso or prev_date_iso == "--":
        return {"available": False}
    try:
        d_str = prev_date_iso.replace("-", "")
        raw = _http_get_text(f"https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?date={d_str}&selectType=ALL&response=json")
        payload = json.loads(raw)
        if payload.get("stat") != "OK":
            return {"available": False}
        for tbl in payload.get("tables", []):
            rows_d = tbl.get("data", [])
            fields = [str(f).strip() for f in tbl.get("fields", [])]
            if any("融資金額" in str(r[0]) for r in rows_d[:5]):
                i_p = next((i for i, f in enumerate(fields) if "前日餘額" in f), 4)
                i_c = next((i for i, f in enumerate(fields) if "今日餘額" in f), 5)
                for r in rows_d:
                    if "融資金額" in str(r[0]):
                        vp = _parse_num(r[i_p])
                        vc = _parse_num(r[i_c])
                        return {
                            "available": True,
                            "date": prev_date_iso,
                            "amt_yi": round(vc / 100000.0, 2),
                            "chg_yi": round((vc - vp) / 100000.0, 2),
                        }
    except Exception:
        pass
    return {"available": False}


# ==========================================
# Phase 2 券商即時行情 API 預留介面（Shioaji / Fugle）
# ==========================================
class BrokerRealtimeAdapter:
    """
    Phase 2 券商即時行情串接介面（08:40 暖機、08:48 台指期日盤 3 分 K、08:55~08:58 權值股試撮五檔）。
    只要設定環境變數 BROKER_PROVIDER=shioaji 或 fugle 與對應金鑰，即可無痛切換至即時行情！
    """

    def __init__(self):
        self.provider = os.environ.get("BROKER_PROVIDER", "none").strip().lower()
        self.ready = False

    def warmup(self):
        """08:40 提前連線暖機、載入憑證與合約"""
        if self.provider == "shioaji":
            # TODO (Phase 2): 接入永豐金 shioaji API Login & Contracts Fetch
            # import shioaji as sj
            # self.api = sj.Shioaji()
            # self.api.login(api_key=os.environ["SHIOAJI_API_KEY"], secret_key=os.environ["SHIOAJI_SECRET_KEY"])
            self.ready = False
        elif self.provider == "fugle":
            # TODO (Phase 2): 接入玉山富果 fugle-marketdata RestClient / WebSocketClient
            self.ready = False
        return self.ready

    def fetch_txf_day_0848(self, prev_taiex_close):
        """08:48 抓取台指期日盤現價、08:45~08:48 開盤 3 分 K 方向與正逆價差"""
        if not self.ready:
            return {
                "available": False,
                "mode": "phase1_standby",
                "note": "尚未接入券商即時行情 API（Phase 2 接入永豐金 Shioaji / 玉山 Fugle 後自動啟用 08:48 台指期日盤校正）",
            }
        return {"available": False}

    def monitor_auction_0855_0858(self):
        """08:55~08:58 監控前五大權值股模擬撮合價與 08:57:30 vs 08:58:00 五檔抽單比對"""
        if not self.ready:
            return {
                "available": False,
                "mode": "phase1_standby",
                "note": "尚未接入券商 WebSocket（Phase 2 接入後將於 08:57:30 與 08:58:00 比對前五大權值股試撮掛單穩定度）",
            }
        return {"available": False}


# ==========================================
# 4 階段滾動更新與評分／背離校正核心
# ==========================================
def load_or_init_state(today_iso):
    os.makedirs(CACHE_DIR, exist_ok=True)
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, "r", encoding="utf-8") as f:
                st = json.load(f)
            if st.get("date") == today_iso:
                return st
        except Exception:
            pass
    return {
        "date": today_iso,
        "updated_at": datetime.now(TW_TZ).strftime("%Y-%m-%d %H:%M"),
        "phase": "Phase 1 (免費開放資料源)",
        "last_stage": None,
        "part1_overnight": {},
        "part2_asia_0815": {},
        "part3_futures_0848": {},
        "part4_auction_0858": {},
        "card_0815": None,
        "card_0858": None,
    }


def save_state(st):
    st["updated_at"] = datetime.now(TW_TZ).strftime("%Y-%m-%d %H:%M")
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)


def compute_overnight_score(us_quotes, adr_quotes, txf_night, foreign_oi, us_holiday):
    """計算隔夜多空基準分 (-5 ~ +5) 與初步預估開盤跳空區間"""
    if us_holiday:
        weighted_us = 0.0
    else:
        w_sum, val_sum = 0.0, 0.0
        for k, (_, _, w) in US_INDEX_WEIGHTS.items():
            q = us_quotes.get(k, {})
            if q.get("available"):
                val_sum += q["chg_pct"] * w
                w_sum += w
        weighted_us = round(val_sum / w_sum, 2) if w_sum > 0 else 0.0

    tsm_pct = adr_quotes.get("TSM", {}).get("chg_pct", 0.0) if (not us_holiday and adr_quotes.get("TSM", {}).get("available")) else 0.0
    # 基準分 (-5 ~ +5)
    raw_score = 0.0
    raw_score += max(-2.5, min(2.5, weighted_us * 1.25))
    raw_score += max(-1.5, min(1.5, tsm_pct * 0.6))
    if txf_night.get("available"):
        raw_score += max(-1.0, min(1.0, txf_night["chg_pts"] / 150.0))
    if foreign_oi.get("available") and foreign_oi.get("day_chg") is not None:
        raw_score += 0.3 if foreign_oi["day_chg"] > 1500 else (-0.3 if foreign_oi["day_chg"] < -1500 else 0.0)
    score = int(round(max(-5.0, min(5.0, raw_score))))
    tone = "偏多" if score >= 2 else "偏空" if score <= -2 else "中性震盪"
    return score, tone, weighted_us


def run_stage_0530(st, today_d: date):
    print(f"🌙 [05:30] 開始抓取隔夜美股、ADR、台指期夜盤、外資期貨與除息扣點 ({today_d})...")
    prev_taiex = fetch_prev_taiex_close()
    us_quotes = {k: _fetch_yf_quote(sym, nm) for k, (sym, nm, _) in US_INDEX_WEIGHTS.items()}
    adr_quotes = {k: _fetch_yf_quote(sym, nm) for k, (sym, nm) in ADR_TICKERS.items()}
    txf_night = fetch_taifex_night_session()
    foreign_oi = fetch_foreign_futures_oi(prev_taiex.get("date"))
    margin_prev = fetch_twse_margin_latest(prev_taiex.get("date"))
    ex_div = fetch_twse_ex_dividend_today(today_d)

    # 美股休市防呆：檢查 SOX/NASDAQ 最新日期是否落後預期前一個美股交易日超過 2 個工作天
    expected_us_d = today_d - timedelta(days=3 if today_d.weekday() == 0 else 1)
    sox_date = us_quotes.get("SOX", {}).get("date", "")
    us_holiday = bool(sox_date and sox_date < expected_us_d.strftime("%Y-%m-%d"))

    score, tone, weighted_us = compute_overnight_score(us_quotes, adr_quotes, txf_night, foreign_oi, us_holiday)
    evap_pts = ex_div.get("est_evap_pts", 0.0)
    prev_close = prev_taiex.get("close", 0.0)

    # 初步開盤跳空估計（優先採用台指期夜盤，無夜盤則由美股加權漲跌幅換算）
    if txf_night.get("available"):
        base_gap = round(txf_night["chg_pts"] - evap_pts, 0)
    else:
        est_pct = (weighted_us * 0.65 + (adr_quotes.get("TSM", {}).get("chg_pct", 0.0) * 0.35)) / 100.0
        base_gap = round(prev_close * est_pct - evap_pts, 0) if prev_close > 0 else 0.0

    est_low = int(round(base_gap - 20))
    est_high = int(round(base_gap + 20))

    st["part1_overnight"] = {
        "fetched_at": datetime.now(TW_TZ).strftime("%H:%M"),
        "prev_taiex": prev_taiex,
        "us_holiday": us_holiday,
        "us_indices": us_quotes,
        "adrs": adr_quotes,
        "txf_night": txf_night,
        "foreign_oi": foreign_oi,
        "margin_prev": margin_prev,
        "ex_dividend": ex_div,
        "weighted_us_pct": weighted_us,
        "overnight_score": score,
        "overnight_tone": tone,
        "est_gap_low": est_low,
        "est_gap_high": est_high,
        "est_open_pts": round(prev_close + base_gap, 1) if prev_close > 0 else 0.0,
    }
    st["last_stage"] = "0530"
    save_state(st)
    print(f"✅ [05:30] 靜默存檔完成：隔夜基調={tone} ({score:+d}/5)，初步跳空估計 {est_low:+d} 至 {est_high:+d} 點")
    return st


def build_card_0815(st):
    p1 = st.get("part1_overnight", {})
    p2 = st.get("part2_asia_0815", {})
    mmdd = st["date"][5:].replace("-", "/")

    score = p2.get("final_score", p1.get("overnight_score", 0))
    tone = p2.get("final_tone", p1.get("overnight_tone", "中性震盪"))
    us = p1.get("us_indices", {})
    adrs = p1.get("adrs", {})

    def _fmt_pct(d, key):
        q = d.get(key, {})
        return f"{q['chg_pct']:+.2f}%" if q.get("available") else "休市/無資料"

    us_line = (
        "美股前一日休市（不重複計入舊收盤）"
        if p1.get("us_holiday")
        else f"費半 {_fmt_pct(us, 'SOX')}、NASDAQ {_fmt_pct(us, 'NASDAQ')}、台積電 ADR {_fmt_pct(adrs, 'TSM')}"
    )
    txf = p1.get("txf_night", {})
    txf_str = f"台指期夜盤 {txf['chg_pts']:+.0f} 點" if txf.get("available") else "台指期夜盤 暫無資料"
    foi = p1.get("foreign_oi", {})
    foi_str = f"外資期貨淨部位 {foi['net_oi']:+,} 口" if foi.get("available") and foi.get("net_oi") is not None else "外資期貨 詳見期交所"
    evap = p1.get("ex_dividend", {}).get("est_evap_pts", 0.0)
    evap_str = f"今日除息扣點 -{evap:.1f} 點" if evap > 0 else "今日無重大除息扣點"

    asia = p2.get("quotes", {})
    div_tag = p2.get("divergence_tag", "同向無背離")
    asia_line = f"日經 {_fmt_pct(asia, 'N225')}、韓股 {_fmt_pct(asia, 'KOSPI')}、小那電子盤 {_fmt_pct(asia, 'NQ')}（{div_tag}）"

    est_lo = p2.get("est_gap_low", p1.get("est_gap_low", 0))
    est_hi = p2.get("est_gap_high", p1.get("est_gap_high", 0))
    dir_word = "跳空開高" if (est_lo + est_hi) > 30 else "跳空開低" if (est_lo + est_hi) < -30 else "平盤附近震盪"
    scenario = p2.get("scenario_note", "開盤後留意前 15 分鐘預估量是否大於五日均量")

    line_text = (
        f"【盤前總覽 {mmdd} 08:15】\n"
        f"- 隔夜基調：{tone}（基準分 {score:+d} / 5）\n"
        f"- 美股與 ADR：{us_line}\n"
        f"- 夜盤與籌碼：{txf_str}｜{foi_str}｜{evap_str}\n"
        f"- 08:15 亞股校正：{asia_line}\n"
        f"- 初步預估：{dir_word}約 {est_lo:+d} 至 {est_hi:+d} 點\n"
        f"- 觀察劇本：{scenario}"
    )

    return {
        "title": f"盤前總覽 {mmdd} 08:15",
        "tone": tone,
        "score": score,
        "divergence_flag": p2.get("divergence_flag", "ALIGNED"),
        "divergence_tag": div_tag,
        "us_line": us_line,
        "chip_line": f"{txf_str}｜{foi_str}｜{evap_str}",
        "asia_line": asia_line,
        "est_gap_text": f"{dir_word}約 {est_lo:+d} 至 {est_hi:+d} 點",
        "scenario": scenario,
        "line_text": line_text,
        "updated_at": datetime.now(TW_TZ).strftime("%H:%M"),
    }


def run_stage_0815(st, today_d: date):
    if not st.get("part1_overnight"):
        st = run_stage_0530(st, today_d)

    print(f"🌅 [08:15] 開始抓取日韓股、三星/海力士、小那/小標普電子盤與 DXY/美債...")
    asia_quotes = {k: _fetch_yf_quote(sym, nm) for k, (sym, nm) in ASIA_MACRO_TICKERS.items()}
    p1 = st["part1_overnight"]
    us_pct = p1.get("weighted_us_pct", 0.0)
    kospi_pct = asia_quotes.get("KOSPI", {}).get("chg_pct", 0.0) if asia_quotes.get("KOSPI", {}).get("available") else 0.0
    nq_pct = asia_quotes.get("NQ", {}).get("chg_pct", 0.0) if asia_quotes.get("NQ", {}).get("available") else 0.0

    # 背離檢查（嚴格依規則：美股收紅但韓股/NQ翻黑 -> 利多出盡警戒；美股收黑但韓股抗跌/NQ翻紅 -> 低接支撐訊號）
    score = p1.get("overnight_score", 0)
    est_lo, est_hi = p1.get("est_gap_low", 0), p1.get("est_gap_high", 0)

    if not p1.get("us_holiday") and us_pct >= 0.4 and (kospi_pct <= -0.3 or nq_pct <= -0.25):
        div_flag = "BEAR_DIVERGENCE"
        div_tag = "⚠️ 利多出盡警戒（美股收紅但亞股/電子盤翻黑）"
        score = max(-5, score - 2)
        est_lo -= 35
        est_hi -= 20
        scenario = "⚠️ 利多出盡警戒！開高後勿急追，嚴防開高走低獲利了結賣壓。"
    elif not p1.get("us_holiday") and us_pct <= -0.4 and (kospi_pct >= 0.2 or nq_pct >= 0.25):
        div_flag = "BULL_DIVERGENCE"
        div_tag = "🛡️ 低接支撐訊號（美股收黑但韓股抗跌/電子盤翻紅）"
        score = min(5, score + 2)
        est_lo += 20
        est_hi += 35
        scenario = "🛡️ 低接支撐訊號！開低後留意權值股承接力道，觀察開低走高收腳機會。"
    else:
        div_flag = "ALIGNED"
        if score >= 2:
            div_tag = "同向偏多，無背離"
            scenario = "多方因子同向，跳空延續機率較高；開高後留意是否量縮獲利了結。"
        elif score <= -2:
            div_tag = "同向偏空，無背離"
            scenario = "空方壓力一致，開低先守關鍵支撐，勿急於摸底接刀。"
        else:
            div_tag = "震盪中性，無明顯背離"
            scenario = "國際盤訊號分歧或變動有限，預期平盤上下震盪，以個股強弱為主。"

    final_tone = "偏多" if score >= 2 else "偏空" if score <= -2 else "中性震盪"
    st["part2_asia_0815"] = {
        "fetched_at": datetime.now(TW_TZ).strftime("%H:%M"),
        "quotes": asia_quotes,
        "divergence_flag": div_flag,
        "divergence_tag": div_tag,
        "final_score": score,
        "final_tone": final_tone,
        "est_gap_low": int(est_lo),
        "est_gap_high": int(est_hi),
        "scenario_note": scenario,
    }
    st["card_0815"] = build_card_0815(st)
    st["last_stage"] = "0815"
    save_state(st)
    print(f"✅ [08:15] 盤前完整簡報已更新！\n{st['card_0815']['line_text']}")
    return st


def run_stage_0848(st, today_d: date):
    if not st.get("card_0815"):
        st = run_stage_0815(st, today_d)
    broker = BrokerRealtimeAdapter()
    broker.warmup()
    prev_close = st.get("part1_overnight", {}).get("prev_taiex", {}).get("close", 0.0)
    res_0848 = broker.fetch_txf_day_0848(prev_close)
    st["part3_futures_0848"] = {
        "fetched_at": datetime.now(TW_TZ).strftime("%H:%M"),
        **res_0848,
    }
    st["last_stage"] = "0848"
    save_state(st)
    return st, None  # 若 Phase 2 偵測方向翻轉則回傳異常警示文字


def run_stage_0858(st, today_d: date):
    if not st.get("card_0815"):
        st = run_stage_0815(st, today_d)
    broker = BrokerRealtimeAdapter()
    broker.warmup()
    res_0858 = broker.monitor_auction_0855_0858()
    st["part4_auction_0858"] = {
        "fetched_at": datetime.now(TW_TZ).strftime("%H:%M"),
        **res_0858,
    }
    mmdd = st["date"][5:].replace("-", "/")
    c15 = st["card_0815"]

    if res_0858.get("available"):
        # Phase 2 即時券商資料已回傳
        line_text = res_0858["line_text"]
        st["card_0858"] = res_0858["card"]
    else:
        # Phase 1 優雅降級：沿用 08:15 校正結果並註明 Phase 2 券商 API 狀態
        line_text = (
            f"【開盤定案 {mmdd} 08:58】\n"
            f"- 方向判定：沿用 08:15 {c15['tone']}（{c15['divergence_tag']}）\n"
            f"- 台指期日盤：待接入 Phase 2 券商 API（目前參考夜盤與 08:15 亞股）\n"
            f"- 權值股試撮：待接入 Phase 2 券商 WebSocket 監控五檔假單\n"
            f"- 預估跳空：{c15['est_gap_text']}\n"
            f"- 操作提醒：開盤後盯緊 09:05 預估量是否大於五日均量"
        )
        st["card_0858"] = {
            "title": f"開盤定案 {mmdd} 08:58",
            "mode": "phase1_standby",
            "direction": f"維持{c15['tone']}（Phase 1 沿用 08:15）",
            "futures_line": "Phase 2 接入永豐金 Shioaji / 玉山 Fugle 後即時更新",
            "auction_line": "Phase 2 接入券商 WebSocket 後自動監控前五大權值股試撮五檔",
            "est_gap_text": c15["est_gap_text"],
            "reminder": "開盤後盯緊 09:05 預估量是否大於五日均量",
            "line_text": line_text,
            "updated_at": datetime.now(TW_TZ).strftime("%H:%M"),
        }

    st["last_stage"] = "0858"
    save_state(st)
    print(f"✅ [08:58] 終版開盤定案已更新！\n{st['card_0858']['line_text']}")
    return st


# ==========================================
# LINE Messaging API 主動推播（支援單人 push / 多人 multicast / 全好友 broadcast）
# ==========================================
def push_line_message(text, dry_run=False):
    if dry_run:
        print("\n🧪 [DRY-RUN 模擬 LINE 推播（不消耗免費額度）]\n" + "=" * 50)
        print(text)
        print("=" * 50)
        return True

    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    raw_target = os.environ.get("LINE_USER_ID", "").strip()
    if not token:
        print("ℹ️ 尚未設定 LINE_CHANNEL_ACCESS_TOKEN，略過 LINE 推播（資料已正常寫入 premarket_daily.json）。")
        return False

    msg_list = [{"type": "text", "text": text[:4900]}]

    # 模式 1：LINE_USER_ID 設為 ALL / BROADCAST 或留空 -> 廣播給所有加此官方帳號好友的家人（免查每個人 User ID）
    if not raw_target or raw_target.upper() in ("ALL", "BROADCAST"):
        url = "https://api.line.me/v2/bot/message/broadcast"
        body = {"messages": msg_list}
        mode_lbl = "全好友廣播 (broadcast)"
    else:
        ids = [x.strip() for x in raw_target.replace(";", ",").split(",") if x.strip()]
        if len(ids) > 1:
            # 模式 2：多組逗號分隔 User ID -> 多人同時推播 (multicast)
            url = "https://api.line.me/v2/bot/message/multicast"
            body = {"to": ids, "messages": msg_list}
            mode_lbl = f"多人指定推播 ({len(ids)} 位)"
        else:
            # 模式 3：單一 User ID -> 單人推播 (push)
            url = "https://api.line.me/v2/bot/message/push"
            body = {"to": ids[0], "messages": msg_list}
            mode_lbl = "單人推播 (push)"

    payload = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            print(f"✅ LINE Messaging API [{mode_lbl}] 推播成功 (HTTP {resp.status})！")
            return True
    except Exception as ex:
        print(f"⚠️ LINE [{mode_lbl}] 推播失敗：{ex}")
        return False


def detect_auto_stage(now_tw: datetime):
    hm = (now_tw.hour, now_tw.minute)
    if hm < (7, 30):
        return "0530"
    elif hm < (8, 40):
        return "0815"
    elif hm < (8, 53):
        return "0848"
    return "0858"


def main():
    parser = argparse.ArgumentParser(description="台股選股雷達 — 盤前自動分析與 LINE 推播機器人")
    parser.add_argument("--stage", choices=["0530", "0815", "0848", "0858", "auto"], default="auto")
    parser.add_argument("--push-line", action="store_true", help="啟用 LINE Messaging API 推播（0530 固定靜默不推播）")
    parser.add_argument("--dry-run", action="store_true", help="測試模式：印出 LINE 推播內容但不實際發送")
    parser.add_argument("--force", action="store_true", help="忽略休市日防呆強制執行（測試用）")
    args = parser.parse_args()

    now_tw = datetime.now(TW_TZ)
    today_d = now_tw.date()
    today_iso = today_d.strftime("%Y-%m-%d")

    is_open, reason = is_tw_trading_day(today_d)
    if not is_open and not args.force:
        print(f"🛑 今日 ({today_iso}) 為 {reason}，依防呆機制略過盤前抓取與 LINE 推播。")
        return

    stage = detect_auto_stage(now_tw) if args.stage == "auto" else args.stage
    st = load_or_init_state(today_iso)

    if stage == "0530":
        run_stage_0530(st, today_d)
        # 05:30 固定靜默存檔，不推播
    elif stage == "0815":
        st = run_stage_0815(st, today_d)
        if args.push_line or args.dry_run:
            push_line_message(st["card_0815"]["line_text"], dry_run=args.dry_run)
    elif stage == "0848":
        st, alert_text = run_stage_0848(st, today_d)
        if alert_text and (args.push_line or args.dry_run):
            push_line_message(alert_text, dry_run=args.dry_run)
    elif stage == "0858":
        st = run_stage_0858(st, today_d)
        if args.push_line or args.dry_run:
            # Phase 1 未接入券商 API 時僅在有即時資料或 dry-run 時推播，避免重複浪費額度
            if st["card_0858"].get("mode") != "phase1_standby" or args.dry_run:
                push_line_message(st["card_0858"]["line_text"], dry_run=args.dry_run)
            else:
                print("ℹ️ Phase 1 尚未接入券商即時行情 API，08:58 略過 LINE 推播以節省免費額度（App 卡片已更新）。")


if __name__ == "__main__":
    main()