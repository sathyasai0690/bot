#!/usr/bin/env python3
"""
BOT3 - Four Indicator Breakout Options Bot
Independent implementation for Zerodha Kite Connect.

IMPORTANT SAFETY DEFAULTS
- LIVE_MODE=False. This is deliberately paper-only until separately reviewed.
- No strategy can guarantee execution prices or profit.
- Uses completed 5-minute underlying candles only.
- Excludes same-day-expiry contracts to avoid expiry-day/physical-settlement risk.
- For NIFTY it selects the nearest future NIFTY expiry; for stocks, the nearest
  future stock-option expiry (monthly under NSE contract specifications).
- Does not infer Greeks. ITM is determined directly from spot and strike.
- Paper fills use the option LTP and are optimistic; reports label them simulated.
"""
from __future__ import annotations

import os
import sys
import json
import time
import logging
import datetime as dt
from pathlib import Path
from typing import Optional

import pandas as pd
from kiteconnect import KiteConnect
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
CFG = {
    "BOT_NAME": "BOT3",
    "LIVE_MODE": False,                 # SAFETY LOCK: do not enable for first run
    "TOTAL_CAPITAL": 50_000.0,
    "MAX_CAPITAL_PER_TRADE": 25_000.0,
    "MAX_OPEN_POSITIONS": 1,
    "TARGET_PROFIT": 1_000.0,            # user’s earlier stated target
    "HARD_STOP_LOSS": 750.0,             # conservative initial risk cap; review before use
    "DAILY_MAX_LOSS": 1_000.0,
    "MAX_TRADES_PER_DAY": 2,
    "ENTRY_START": dt.time(9, 30),
    "ENTRY_END": dt.time(14, 30),
    "SQUARE_OFF": dt.time(15, 10),
    "CANDLE_INTERVAL": "5minute",
    "SUPERTREND_PERIOD": 10,
    "SUPERTREND_MULTIPLIER": 3.0,
    "RSI_PERIOD": 14,
    "RSI_CE_MIN": 65.0,
    "RSI_PE_MAX": 35.0,
    "BB_PERIOD": 20,
    "BB_STD": 2.0,
    "OPTION_MIN_LTP": 10.0,
    "OPTION_MAX_LTP": 500.0,
    "MIN_OPTION_OI": 0,                 # set a meaningful broker-verified threshold if desired
    "MAX_SPREAD_PCT": 3.0,
    "MIN_DTE": 1,                       # exclude expiry day (0 DTE)
    "POLL_SECONDS": 5,
    "SCAN_CYCLE_SECONDS": 300,        # evaluate signals once per completed 5m bar
    "LOG_DIR": "bot3_logs",
    "OUTPUT_DIR": "bot3_outputs",
    "TOKEN_FILE": "kite_token.json",
    "CREDS_FILE": "kite_creds.json",
    # Uses the previous completed daily candle for traditional pivots.
    "UNIVERSE_SOURCE": "live_nfo_instrument_master",
}

Path(CFG["LOG_DIR"]).mkdir(exist_ok=True)
Path(CFG["OUTPUT_DIR"]).mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | BOT3 | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(Path(CFG["LOG_DIR"]) / "bot3.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("BOT3")

TODAY = dt.datetime.now(IST).date()
REPORT = Path(CFG["OUTPUT_DIR"]) / f"bot3_report_{TODAY.isoformat()}.xlsx"
STATE = {
    "kite": None, "universe": [], "positions": {}, "trades": [],
    "decisions": [], "seen_candles": {}, "realized_pnl": 0.0,
    "trade_count": 0, "instrument_map": None, "halted": False,
    "last_scan_bucket": None,
}


def now_ist():
    return dt.datetime.now(IST)


def kite_connect():
    if not Path(CFG["CREDS_FILE"]).exists():
        raise RuntimeError(f"Missing {CFG['CREDS_FILE']} with api_key and api_secret")
    creds = json.loads(Path(CFG["CREDS_FILE"]).read_text())
    kite = KiteConnect(api_key=creds["api_key"])
    token_path = Path(CFG["TOKEN_FILE"])
    if token_path.exists():
        token_data = json.loads(token_path.read_text())
        if token_data.get("date") == TODAY.isoformat():
            kite.set_access_token(token_data["access_token"])
            kite.profile()
            return kite
    print(kite.login_url())
    request_token = input("Paste Kite request_token: ").strip()
    session = kite.generate_session(request_token, api_secret=creds["api_secret"])
    kite.set_access_token(session["access_token"])
    token_path.write_text(json.dumps({
        "date": TODAY.isoformat(), "access_token": session["access_token"]
    }))
    kite.profile()
    return kite


def instrument_master(kite):
    # Fetch once; all contract choices are validated against this live master.
    if STATE["instrument_map"] is None:
        STATE["instrument_map"] = kite.instruments("NFO")
    return STATE["instrument_map"]


def build_universe(kite):
    nfo = instrument_master(kite)
    option_underlyings = {
        x.get("name") for x in nfo
        if x.get("segment") == "NFO-OPT"
        and x.get("instrument_type") in ("CE", "PE")
    }
    # Underlying NSE EQ tokens; NIFTY index token is separately included.
    eq = {x["tradingsymbol"]: x for x in kite.instruments("NSE")
          if x.get("segment") == "NSE" and x.get("instrument_type") == "EQ"}
    symbols = sorted(s for s in option_underlyings if s in eq)
    # "Affordable" is enforced at option-lot premium cost at contract selection,
    # not by share price alone. NIFTY is always added.
    universe = [{"name": s, "token": eq[s]["instrument_token"], "kind": "STOCK"}
                for s in symbols]
    universe.append({"name": "NIFTY", "token": 256265, "kind": "INDEX"})
    log.info("Loaded %d live F&O stock underlyings plus NIFTY", len(symbols))
    return universe


def fetch_history(kite, token, interval="5minute", days=12):
    end = now_ist().replace(tzinfo=None)
    start = end - dt.timedelta(days=days)
    bars = kite.historical_data(token, start, end, interval)
    if not bars:
        return pd.DataFrame()
    df = pd.DataFrame(bars)
    df["date"] = pd.to_datetime(df["date"])
    if df["date"].dt.tz is None:
        df["date"] = df["date"].dt.tz_localize("Asia/Kolkata")
    else:
        df["date"] = df["date"].dt.tz_convert("Asia/Kolkata")
    df = df.set_index("date").sort_index()
    # Drop currently forming candle. Only completed bars are eligible.
    current = now_ist()
    if len(df):
        last = df.index[-1]
        if (current - last).total_seconds() < 300:
            df = df.iloc[:-1]
    return df


def rsi(series, period=14):
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, float("nan"))
    out = 100 - (100 / (1 + rs))
    out = out.mask((avg_loss == 0) & (avg_gain > 0), 100)
    out = out.mask((avg_gain == 0) & (avg_loss > 0), 0)
    return out


def supertrend(df, period=10, multiplier=3.0):
    high, low, close = df["high"], df["low"], df["close"]
    prev_close = close.shift(1)
    tr = pd.concat([
        high - low, (high - prev_close).abs(), (low - prev_close).abs()
    ], axis=1).max(axis=1)
    atr = tr.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    mid = (high + low) / 2
    upper = mid + multiplier * atr
    lower = mid - multiplier * atr
    final_upper = upper.copy()
    final_lower = lower.copy()
    st = pd.Series(index=df.index, dtype=float)
    direction = pd.Series(index=df.index, dtype=bool)
    for i in range(len(df)):
        if pd.isna(atr.iloc[i]):
            continue
        if i == 0 or pd.isna(st.iloc[i-1]):
            direction.iloc[i] = True
            st.iloc[i] = lower.iloc[i]
            continue
        final_upper.iloc[i] = (
            upper.iloc[i] if upper.iloc[i] < final_upper.iloc[i-1]
            or close.iloc[i-1] > final_upper.iloc[i-1] else final_upper.iloc[i-1]
        )
        final_lower.iloc[i] = (
            lower.iloc[i] if lower.iloc[i] > final_lower.iloc[i-1]
            or close.iloc[i-1] < final_lower.iloc[i-1] else final_lower.iloc[i-1]
        )
        if st.iloc[i-1] == final_upper.iloc[i-1]:
            direction.iloc[i] = close.iloc[i] > final_upper.iloc[i]
        else:
            direction.iloc[i] = close.iloc[i] >= final_lower.iloc[i]
        st.iloc[i] = final_lower.iloc[i] if direction.iloc[i] else final_upper.iloc[i]
    return st, direction


def traditional_pivots(kite, token):
    daily = fetch_history(kite, token, interval="day", days=12)
    if len(daily) < 2:
        return None
    # Previous fully completed daily session, not today's developing candle.
    prev = daily.iloc[-1]
    h, l, c = float(prev["high"]), float(prev["low"]), float(prev["close"])
    p = (h + l + c) / 3.0
    return {"P": p, "R1": 2*p-l, "S1": 2*p-h, "source_date": str(daily.index[-1].date())}


def make_signal(kite, item):
    df = fetch_history(kite, item["token"], days=12)
    if len(df) < 40:
        return None, "Insufficient completed 5m candles"
    close = df["close"]
    st, st_up = supertrend(df, CFG["SUPERTREND_PERIOD"], CFG["SUPERTREND_MULTIPLIER"])
    rv = rsi(close, CFG["RSI_PERIOD"])
    mid = close.rolling(CFG["BB_PERIOD"]).mean()
    sd = close.rolling(CFG["BB_PERIOD"]).std(ddof=0)
    upper, lower = mid + CFG["BB_STD"]*sd, mid - CFG["BB_STD"]*sd
    piv = traditional_pivots(kite, item["token"])
    if piv is None:
        return None, "Previous-day pivot unavailable"
    bar = df.iloc[-1]
    close_px = float(bar["close"])
    values = {
        "candle_time": str(df.index[-1]),
        "underlying_close": close_px,
        "supertrend": float(st.iloc[-1]) if pd.notna(st.iloc[-1]) else None,
        "supertrend_up": bool(st_up.iloc[-1]) if pd.notna(st.iloc[-1]) else None,
        "rsi14": float(rv.iloc[-1]) if pd.notna(rv.iloc[-1]) else None,
        "bb_upper": float(upper.iloc[-1]) if pd.notna(upper.iloc[-1]) else None,
        "bb_lower": float(lower.iloc[-1]) if pd.notna(lower.iloc[-1]) else None,
        "pivot_R1": piv["R1"], "pivot_S1": piv["S1"],
        "pivot_source_date": piv["source_date"],
    }
    if any(values[k] is None for k in ("supertrend", "rsi14", "bb_upper", "bb_lower")):
        return None, "Indicator warm-up incomplete"
    ce = [
        close_px > values["supertrend"] and values["supertrend_up"],
        values["rsi14"] > CFG["RSI_CE_MIN"],
        close_px > piv["R1"],
        close_px > values["bb_upper"],
    ]
    pe = [
        close_px < values["supertrend"] and not values["supertrend_up"],
        values["rsi14"] < CFG["RSI_PE_MAX"],
        close_px < piv["S1"],
        close_px < values["bb_lower"],
    ]
    values["CE_pass"] = ce
    values["PE_pass"] = pe
    values["decision"] = "CE" if all(ce) else ("PE" if all(pe) else "NO_TRADE")
    return values, "All four conditions passed" if values["decision"] != "NO_TRADE" else "One or more conditions failed"


def choose_option(kite, underlying, direction, spot):
    today = now_ist().date()
    insts = instrument_master(kite)
    candidates = [
        x for x in insts
        if x.get("name") == underlying
        and x.get("instrument_type") == direction
        and x.get("expiry") and (x["expiry"] - today).days >= CFG["MIN_DTE"]
        and x.get("segment") == "NFO-OPT"
    ]
    if not candidates:
        return None, "No non-expiry-day option contract in instrument master"
    expiry = min(x["expiry"] for x in candidates)
    candidates = [x for x in candidates if x["expiry"] == expiry]
    # ITM only: CE strike < spot; PE strike > spot. Exclude ATM/OTM.
    if direction == "CE":
        candidates = [x for x in candidates if float(x["strike"]) < spot]
    else:
        candidates = [x for x in candidates if float(x["strike"]) > spot]
    if not candidates:
        return None, "No strictly ITM strike available"
    candidates.sort(key=lambda x: abs(spot - float(x["strike"])))
    for inst in candidates[:20]:
        key = f"NFO:{inst['tradingsymbol']}"
        try:
            q = kite.quote([key])[key]
            ltp = float(q.get("last_price") or 0)
            depth = q.get("depth", {})
            buy = (depth.get("buy") or [{}])[0].get("price", 0)
            sell = (depth.get("sell") or [{}])[0].get("price", 0)
            oi = int(q.get("oi") or 0)
            spread_pct = ((sell-buy)/((sell+buy)/2)*100) if buy and sell and sell >= buy else 999
            lot = int(inst["lot_size"])
            cost = ltp * lot
            if not CFG["OPTION_MIN_LTP"] <= ltp <= CFG["OPTION_MAX_LTP"]:
                continue
            if CFG["MIN_OPTION_OI"] and oi < CFG["MIN_OPTION_OI"]:
                continue
            if spread_pct > CFG["MAX_SPREAD_PCT"]:
                continue
            if cost > CFG["MAX_CAPITAL_PER_TRADE"] or cost > CFG["TOTAL_CAPITAL"]:
                continue
            return {
                "symbol": inst["tradingsymbol"], "token": inst["instrument_token"],
                "strike": float(inst["strike"]), "expiry": str(expiry),
                "lot_size": lot, "ltp": ltp, "oi": oi,
                "spread_pct": round(spread_pct, 3), "estimated_lot_cost": cost,
            }, "Selected nearest liquid strictly ITM contract"
        except Exception as exc:
            log.warning("Quote failed for %s: %s", key, exc)
    return None, "No ITM option passed premium, spread and capital filters"


def append_decision(record):
    STATE["decisions"].append(record)
    write_report()


def write_report():
    headers = {
        "Summary": ["Metric", "Value"],
        "Decisions": ["Timestamp", "Underlying", "Kind", "Candle", "Close", "ST", "ST Up",
                      "RSI14", "R1", "S1", "BB Upper", "BB Lower", "CE checks", "PE checks",
                      "Decision", "Reason", "Option", "Strike", "Expiry", "Option LTP"],
        "Trades": ["Entry time", "Underlying", "Direction", "Option", "Strike", "Expiry",
                   "Qty", "Entry premium", "Exit time", "Exit premium", "Gross P&L",
                   "Estimated charges", "Net P&L", "Exit reason", "Execution mode"],
    }
    if REPORT.exists():
        wb = load_workbook(REPORT)
    else:
        wb = Workbook()
        wb.remove(wb.active)
        for name, hdr in headers.items():
            ws = wb.create_sheet(name)
            ws.append(hdr)
    for name, hdr in headers.items():
        if name not in wb.sheetnames:
            ws = wb.create_sheet(name)
            ws.append(hdr)
    ws = wb["Summary"]
    ws.delete_rows(1, ws.max_row)
    summary = [
        ("Date", TODAY.isoformat()), ("Mode", "LIVE" if CFG["LIVE_MODE"] else "PAPER"),
        ("Capital", CFG["TOTAL_CAPITAL"]), ("Target per trade", CFG["TARGET_PROFIT"]),
        ("Hard stop per trade", CFG["HARD_STOP_LOSS"]),
        ("Closed trades", len(STATE["trades"])),
        ("Realized net P&L", round(sum(t.get("net_pnl", 0) for t in STATE["trades"]), 2)),
        ("Open positions", len(STATE["positions"])),
        ("Entry rules", "Completed 5m candle; ST(10,3), RSI14 >65/<35, close >R1/<S1, close outside BB(20,2)"),
        ("Expiry rule", "NIFTY nearest future expiry; stocks nearest future monthly expiry; same-day expiry excluded"),
        ("Fill model", "Paper mode uses LTP; simulated fills are optimistic and exclude market impact"),
    ]
    ws.append(["Metric", "Value"])
    for row in summary: ws.append(list(row))
    ws = wb["Decisions"]
    ws.delete_rows(2, ws.max_row)
    for d in STATE["decisions"]:
        s = d.get("signal") or {}
        o = d.get("option") or {}
        ws.append([
            d["timestamp"], d["underlying"], d["kind"], s.get("candle_time"),
            s.get("underlying_close"), s.get("supertrend"), s.get("supertrend_up"),
            s.get("rsi14"), s.get("pivot_R1"), s.get("pivot_S1"), s.get("bb_upper"),
            s.get("bb_lower"), str(s.get("CE_pass")), str(s.get("PE_pass")),
            s.get("decision"), d.get("reason"), o.get("symbol"), o.get("strike"),
            o.get("expiry"), o.get("ltp")
        ])
    ws = wb["Trades"]
    ws.delete_rows(2, ws.max_row)
    for t in STATE["trades"]:
        ws.append([t.get(k) for k in (
            "entry_time", "underlying", "direction", "option", "strike", "expiry",
            "qty", "entry_price", "exit_time", "exit_price", "gross_pnl",
            "estimated_charges", "net_pnl", "exit_reason", "mode"
        )])
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
            cell.alignment = Alignment(wrap_text=True)
        for col in sheet.columns:
            letter = col[0].column_letter
            sheet.column_dimensions[letter].width = min(
                36, max(12, max(len(str(c.value or "")) for c in col[:100]) + 2)
            )
    wb.save(REPORT)


def enter_position(kite, item, signal, opt):
    qty = opt["lot_size"]  # one lot only for risk containment
    cost = qty * opt["ltp"]
    if cost > CFG["MAX_CAPITAL_PER_TRADE"] or cost > CFG["TOTAL_CAPITAL"]:
        return False, "Lot cost exceeds capital limit"
    if CFG["LIVE_MODE"]:
        # Deliberately disabled: live order execution needs order-state reconciliation,
        # partial-fill handling, broker RMS checks and explicit operator approval.
        return False, "LIVE order path is locked pending execution-safety validation"
    pos_id = opt["symbol"]
    if pos_id in STATE["positions"]:
        return False, "Position already exists"
    STATE["positions"][pos_id] = {
        "id": pos_id, "underlying": item["name"], "kind": item["kind"],
        "direction": signal["decision"], "option": opt["symbol"],
        "strike": opt["strike"], "expiry": opt["expiry"], "qty": qty,
        "entry_price": opt["ltp"], "entry_time": now_ist(),
        "peak_pnl": 0.0, "trail_floor": None, "token": opt["token"],
    }
    STATE["trade_count"] += 1
    log.info("PAPER ENTRY %s %s qty=%s premium=%.2f cost=%.2f",
             item["name"], opt["symbol"], qty, opt["ltp"], cost)
    return True, "Paper entry recorded"


def close_position(kite, pos, ltp, reason):
    gross = (ltp - pos["entry_price"]) * pos["qty"]
    # Charges are not fabricated; user should reconcile exact broker contract-note costs.
    charges = 0.0
    net = gross - charges
    STATE["realized_pnl"] += net
    STATE["trades"].append({
        "entry_time": pos["entry_time"].isoformat(), "underlying": pos["underlying"],
        "direction": pos["direction"], "option": pos["option"], "strike": pos["strike"],
        "expiry": pos["expiry"], "qty": pos["qty"], "entry_price": pos["entry_price"],
        "exit_time": now_ist().isoformat(), "exit_price": ltp, "gross_pnl": gross,
        "estimated_charges": charges, "net_pnl": net, "exit_reason": reason,
        "mode": "PAPER_SIMULATED",
    })
    STATE["positions"].pop(pos["id"], None)
    log.info("PAPER EXIT %s P&L=%.2f reason=%s", pos["option"], net, reason)
    write_report()


def monitor_positions(kite):
    for pos in list(STATE["positions"].values()):
        try:
            q = kite.quote([f"NFO:{pos['option']}"])[f"NFO:{pos['option']}"]
            ltp = float(q.get("last_price") or 0)
            if ltp <= 0: continue
            pnl = (ltp - pos["entry_price"]) * pos["qty"]
            pos["peak_pnl"] = max(pos["peak_pnl"], pnl)
            if pos["peak_pnl"] >= CFG["TARGET_PROFIT"]:
                pos["trail_floor"] = max(0.0, pos["peak_pnl"] - 250.0)
            reason = None
            if pnl <= -CFG["HARD_STOP_LOSS"]:
                reason = "Hard stop loss"
            elif pos["trail_floor"] is not None and pnl <= pos["trail_floor"]:
                reason = "Trailing floor"
            elif now_ist().time() >= CFG["SQUARE_OFF"]:
                reason = "Intraday square-off"
            if reason:
                # Paper exit at observed LTP only; no broker order is sent.
                close_position(kite, pos, ltp, reason)
        except Exception as exc:
            log.exception("Position monitor error %s: %s", pos["option"], exc)


def scan_once(kite):
    now = now_ist()
    monitor_positions(kite)
    if STATE["halted"] or now.time() < CFG["ENTRY_START"] or now.time() > CFG["ENTRY_END"]:
        return
    # Avoid repeatedly requesting the same 5-minute candle for every symbol.
    bucket = int(now.timestamp() // CFG["SCAN_CYCLE_SECONDS"])
    if STATE["last_scan_bucket"] == bucket:
        return
    STATE["last_scan_bucket"] = bucket
    if STATE["realized_pnl"] <= -CFG["DAILY_MAX_LOSS"]:
        STATE["halted"] = True
        log.error("Daily loss limit reached; new entries halted")
        return
    if STATE["trade_count"] >= CFG["MAX_TRADES_PER_DAY"]:
        return
    if len(STATE["positions"]) >= CFG["MAX_OPEN_POSITIONS"]:
        return
    for item in STATE["universe"]:
        try:
            signal, reason = make_signal(kite, item)
            if signal is None:
                continue
            candle_key = f"{item['name']}|{signal['candle_time']}"
            if STATE["seen_candles"].get(item["name"]) == signal["candle_time"]:
                continue
            STATE["seen_candles"][item["name"]] = signal["candle_time"]
            record = {
                "timestamp": now.isoformat(), "underlying": item["name"],
                "kind": item["kind"], "signal": signal, "reason": reason, "option": None,
            }
            if signal["decision"] == "NO_TRADE":
                append_decision(record)
                continue
            spot = signal["underlying_close"]
            opt, opt_reason = choose_option(kite, item["name"], signal["decision"], spot)
            record["option"] = opt
            record["reason"] = opt_reason
            append_decision(record)
            if not opt:
                continue
            ok, entry_reason = enter_position(kite, item, signal, opt)
            log.info("%s %s: %s", item["name"], signal["decision"], entry_reason)
            if ok:
                write_report()
                return  # one new entry per scan cycle
        except Exception as exc:
            log.exception("Scan failed for %s: %s", item["name"], exc)


def main():
    log.info("Starting %s | %s | capital Rs %.2f",
             CFG["BOT_NAME"], "LIVE LOCKED" if CFG["LIVE_MODE"] else "PAPER",
             CFG["TOTAL_CAPITAL"])
    if CFG["LIVE_MODE"]:
        raise RuntimeError("LIVE_MODE is not supported in this build; safety gate prevents order placement.")
    kite = kite_connect()
    STATE["kite"] = kite
    STATE["universe"] = build_universe(kite)
    if not STATE["universe"]:
        raise RuntimeError("No eligible F&O underlyings found; aborting.")
    log.info("Universe ready. Tomorrow-specific contracts will be selected from live NFO master.")
    write_report()
    while True:
        try:
            if now_ist().date() != TODAY:
                log.info("Session date changed; stopping bot.")
                break
            scan_once(kite)
            time.sleep(CFG["POLL_SECONDS"])
        except KeyboardInterrupt:
            log.info("Keyboard interrupt; stopping scan loop.")
            break
        except Exception:
            log.exception("Main loop error; retrying after 5 seconds.")
            time.sleep(5)
    # Final report. Paper positions are not carried overnight.
    write_report()
    log.info("BOT3 stopped. Report: %s", REPORT)


if __name__ == "__main__":
    main()
