# -*- coding: utf-8 -*-
"""
每日 15:30 自動預抓選股快取腳本 (搭配 GitHub Actions 排程執行)
功能：每天週一至週五台灣時間 15:35 自動掃描「熱門權值股、上市股、上櫃股、興櫃股」並補齊三大法人籌碼成本，
     將結果輸出為 daily_cache/<mode>.json 供手機網頁版 app.py 0.1 秒極速秒載，完全避免多人使用造成卡頓或限流！
"""

import os
import json
import time
from datetime import datetime, timezone, timedelta

# 為了讓此獨立腳本可直接複用 app.py 內的運算邏輯，先建立一個輕量假 streamlit 模組以略過 UI 執行
import sys
import types


class _DummyCtx:
    def __enter__(self): return self
    def __exit__(self, *args): pass


class _DummySt(types.ModuleType):
    secrets = {}
    session_state = {}
    query_params = {}
    def set_page_config(self, *a, **k): pass
    def markdown(self, *a, **k): pass
    def title(self, *a, **k): pass
    def info(self, *a, **k): pass
    def success(self, *a, **k): pass
    def warning(self, *a, **k): pass
    def error(self, *a, **k): pass
    def caption(self, *a, **k): pass
    def subheader(self, *a, **k): pass
    def write(self, *a, **k): pass
    def altair_chart(self, *a, **k): pass
    def rerun(self, *a, **k): pass
    def stop(self): raise SystemExit(0)
    def cache_data(self, *a, **k):
        def _deco(fn): return fn
        return _deco
    def expander(self, *a, **k): return _DummyCtx()
    def form(self, *a, **k): return _DummyCtx()
    def form_submit_button(self, *a, **k): return False
    def columns(self, n):
        cnt = len(n) if isinstance(n, (list, tuple)) else int(n)
        return [self] * cnt
    def number_input(self, *a, **k): return k.get("value", 0)
    def tabs(self, names): return [_DummyCtx() for _ in names]
    def text_input(self, *a, **k): return ""
    def button(self, *a, **k): return False
    def radio(self, *a, **k): return "top"
    def download_button(self, *a, **k): pass
    def spinner(self, *a, **k): return _DummyCtx()
    def metric(self, *a, **k): pass
    def __getattr__(self, name):
        return lambda *a, **k: None


sys.modules["streamlit"] = _DummySt("streamlit")

import yfinance as yf
import pandas as pd
import app as core


def run_precompute():
    now_tw = datetime.now(timezone(timedelta(hours=8)))
    print(f"🚀 開始執行每日 15:30 主動預抓選股快取（台灣時間：{now_tw.strftime('%Y-%m-%d %H:%M')}）")

    _, _, period_key, _ = core.get_market_data_cutoff_info()
    effective_date = core.get_effective_trading_date_str(period_key)
    print(f"📅 偵測到最新有效收盤資料日期：{effective_date}")

    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_cache")
    os.makedirs(cache_dir, exist_ok=True)

    for mode in ("top", "twse", "tpex", "emerging"):
        t0 = time.time()
        min_amt = 500.0 if mode == "emerging" else 2000.0
        scan_params = {
            "fast_len": 12, "slow_len": 26, "sig_len": 9, "macd_tolerance": 3,
            "bb_len": 20, "bb_k": 2.0, "max_pre_bw": 15.0,
            "ma_len": 20, "ma_up_bars": 1, "max_bias20": 8.0,
            "min_amount_wan": min_amt, "vol_ratio": 1.5, "upper_shadow_ratio": 0.5,
        }

        print(f"\n--- 正在預抓【{core.MODE_LABELS[mode]}】(最低成交額 {min_amt} 萬元) ---")
        pkg = core.run_shared_market_scan_cached(mode, period_key, tuple(sorted(scan_params.items())))

        # 清理不可序列化或過大的內部欄位，保留所有畫面與 Excel 所需欄位
        def _clean_rows(rows):
            cleaned = []
            for r in rows:
                c = {k: v for k, v in r.items() if k not in ("_recent_20_prices", "_exchange")}
                cleaned.append(c)
            return cleaned

        out_payload = {
            "mode": mode,
            "data_date": pkg.get("data_date") or effective_date,
            "generated_at": datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M"),
            "valid_cnt": pkg["valid_cnt"],
            "total": pkg["total"],
            "full_sorted": _clean_rows(pkg["full_sorted"]),
            "near_top": _clean_rows(pkg["near_top"]),
            "source": "daily_precomputed",
        }

        out_path = os.path.join(cache_dir, f"{mode}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out_payload, f, ensure_ascii=False, indent=2)

        print(
            f"✅ 【{core.MODE_LABELS[mode]}】預抓完成（耗時 {time.time()-t0:.1f} 秒）："
            f"完全符合 {len(out_payload['full_sorted'])} 檔，準突破 {len(out_payload['near_top'])} 檔 -> 已存至 daily_cache/{mode}.json"
        )

    # 同步維護近 30 天曾上榜股票之連續價量歷史檔 (daily_cache/watchlist_history.json)
    # 讓深色 PWA 自選股即使幾天後不在當日選股榜上，也能天天看到最新股價與成交量變化！
    hist_path = os.path.join(cache_dir, "watchlist_history.json")
    hist_map = {}
    if os.path.exists(hist_path):
        try:
            with open(hist_path, "r", encoding="utf-8") as f:
                hist_map = json.load(f)
        except Exception:
            hist_map = {}

    # 收集今日所有上榜標的（含既有的歷史追蹤代號，保留最多 120 檔避免檔案過大）
    tracked_syms = {}
    for m_key in ("top", "twse", "tpex", "emerging"):
        mp = os.path.join(cache_dir, f"{m_key}.json")
        if os.path.exists(mp):
            with open(mp, "r", encoding="utf-8") as f:
                m_pkg = json.load(f)
            for row in m_pkg.get("full_sorted", []) + m_pkg.get("near_top", []):
                code = row["股票代碼"]
                suf = ".TW" if "上市" in row.get("市場別", "") else ".TWO"
                tracked_syms[code] = (f"{code}{suf}", row["股票名稱"], row.get("_chart_30d"))

    for old_code, old_entry in list(hist_map.items()):
        if old_code not in tracked_syms and len(tracked_syms) < 120:
            tracked_syms[old_code] = (old_entry.get("yf_sym", f"{old_code}.TW"), old_entry.get("name", ""), None)

    need_fetch = [ (code, info[0], info[1]) for code, info in tracked_syms.items() if info[2] is None ]
    for code, (sym, nm, chart_d) in tracked_syms.items():
        if chart_d:
            hist_map[code] = {"yf_sym": sym, "name": nm, "updated": effective_date, "chart": chart_d}

    if need_fetch:
        print(f"\n📈 正在更新 {len(need_fetch)} 檔歷史曾上榜自選追蹤股之最新價量走勢...")
        syms_to_dl = [x[1] for x in need_fetch]
        try:
            dl_data = yf.download(tickers=syms_to_dl, period="3mo", interval="1d", group_by="ticker", auto_adjust=False, threads=True, progress=False)
            for code, sym, nm in need_fetch:
                try:
                    sub = dl_data if len(syms_to_dl) == 1 else dl_data[sym]
                    if sub is None or sub.dropna(how="all").empty:
                        continue
                    sub = sub.dropna(subset=["Open", "Close", "Volume"]).copy()
                    sub = sub[sub["Volume"] > 0]
                    if len(sub) < 5:
                        continue
                    c_s = pd.Series(sub["Close"].squeeze()).astype(float)
                    o_s = pd.Series(sub["Open"].squeeze()).astype(float)
                    v_s = pd.Series(sub["Volume"].squeeze()).astype(float) / 1000.0
                    ma20_s = c_s.rolling(20).mean()
                    vma5_s = v_s.rolling(5).mean()
                    n_p = min(30, len(sub))
                    hist_map[code] = {
                        "yf_sym": sym,
                        "name": nm,
                        "updated": effective_date,
                        "chart": {
                            "dates": [dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10] for dt in sub.index[-n_p:]],
                            "close": [round(float(x), 2) for x in c_s.iloc[-n_p:]],
                            "ma20": [round(float(x), 2) if not pd.isna(x) else None for x in ma20_s.iloc[-n_p:]],
                            "vol": [int(round(float(x))) for x in v_s.iloc[-n_p:]],
                            "vol_ma5": [int(round(float(x))) if not pd.isna(x) else None for x in vma5_s.iloc[-n_p:]],
                            "up": [bool(float(c) >= float(o)) for c, o in zip(c_s.iloc[-n_p:], o_s.iloc[-n_p:])],
                        },
                    }
                except Exception:
                    continue
        except Exception as ex:
            print(f"⚠️ 更新歷史追蹤檔時發生異常：{ex}")

    with open(hist_path, "w", encoding="utf-8") as f:
        json.dump(hist_map, f, ensure_ascii=False)
    print(f"✅ 已同步更新 daily_cache/watchlist_history.json（共追蹤 {len(hist_map)} 檔）！")
    print("\n🎉 所有市場每日預抓快取已完成！")


if __name__ == "__main__":
    run_precompute()