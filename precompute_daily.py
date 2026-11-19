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
    def set_page_config(self, *a, **k): pass
    def markdown(self, *a, **k): pass
    def title(self, *a, **k): pass
    def info(self, *a, **k): pass
    def success(self, *a, **k): pass
    def warning(self, *a, **k): pass
    def error(self, *a, **k): pass
    def caption(self, *a, **k): pass
    def subheader(self, *a, **k): pass
    def stop(self): raise SystemExit(0)
    def cache_data(self, *a, **k):
        def _deco(fn): return fn
        return _deco
    def expander(self, *a, **k): return _DummyCtx()
    def columns(self, n): return [self] * n
    def number_input(self, *a, **k): return k.get("value", 0)
    def tabs(self, names): return [_DummyCtx() for _ in names]
    def text_input(self, *a, **k): return ""
    def button(self, *a, **k): return False
    def radio(self, *a, **k): return "top"
    def download_button(self, *a, **k): pass
    def spinner(self, *a, **k): return _DummyCtx()
    def metric(self, *a, **k): pass


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

    print("\n🎉 所有市場每日預抓快取已完成！")


if __name__ == "__main__":
    run_precompute()