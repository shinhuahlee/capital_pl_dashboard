"""
一次性/可重複執行 的回填腳本：把 docs/history.json 每一天的 tf_daily_pl_equity / of_daily_pl_equity
(權益差額法)算出來並寫回去。

出入金資料讀取來源(來源: 群益網路下單網頁匯出，副檔名 .xls 但實際是 HTML table)：
    國內期貨出入金.xls
    海外期貨出入金.xls
(這兩個檔案含真實交易明細，已加進 .gitignore，不會被 commit)

這兩份 xls 每次下載都只會有「最近一段期間」的出入金，較舊的紀錄可能不會再出現。
所以每次執行都會把新解析到的交易「併入」持久化的 deposit_withdrawal_log.json
(用 市場+日期+時間+金額+類型 當作唯一 key 去重)，不會因為 xls 只剩最近資料就把舊紀錄洗掉。
之後所有計算都是根據這份持久化 log，而不是直接讀當下的 xls。

日損益(權益法) = 今日權益 - 昨日權益 - 當日存提款(出金為負值會被加回，入金為正值會被扣回)

equity 為 0 代表當天 tf_equity/of_equity(權益)、tf_float_pl/of_float_pl(浮動損益)
還沒抓到資料(不是查詢出錯，是當時程式版本還沒抓這兩個欄位)。這種情況下 daily_pl 直接退回
用 tf_total_pl/of_total_pl(close_pl+float_pl) 當作那天的日損益。
同時仍會維護一個「合成權益」往前推進：
    合成權益 += 當天 total_pl + 當天存提款
一旦遇到有真實 equity 的那天，改用真實值跟「合成權益」比較算出當天的 daily_pl(權益差額法)，
合成權益也重置成真實值，避免長期缺值期間的估計誤差持續累積。

起始權益 = BASE_EQUITY_BEFORE_RECORDS(出入金紀錄最早一筆之前，帳戶原本就有的餘額，預設 0)
           + 所有日期早於 history.json 最早一天的出入金加總。
國內帳戶最早紀錄是 2026-04-14 入金 200萬，早於 history 最早的 2026-04-15，所以會自動變成起始權益。
海外帳戶最早紀錄是 2026-05-13(60萬轉入)，晚於 history 最早日期，所以起始權益維持 0，
60萬 改記錄成 2026-05-13 當天的存款。
換匯(台幣↔美元)不算存提款，視為帳戶內部轉換，不影響權益總值，直接跳過不計入。
"已取消"的申請不計入。
"""
import json
import os

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_PATH = os.path.join(BASE_DIR, "docs", "history.json")
TF_XLS_PATH = os.path.join(BASE_DIR, "國內期貨出入金.xls")
OF_XLS_PATH = os.path.join(BASE_DIR, "海外期貨出入金.xls")
DW_LOG_PATH = os.path.join(BASE_DIR, "deposit_withdrawal_log.json")

BASE_EQUITY_BEFORE_RECORDS = {"TF": 0.0, "OF": 0.0}


def _f(val):
    try:
        return float(str(val).replace(",", ""))
    except Exception:
        return 0.0


def parse_tf_transactions():
    df = pd.read_html(TF_XLS_PATH, encoding="utf-8")[0]
    df = df[df["交易日期"] != "總計"].copy()
    df = df[~df["申請狀態"].astype(str).str.contains("取消")]
    df["date"] = df["交易日期"].str.replace("/", "-")
    df["signed"] = df.apply(
        lambda r: _f(r["交易金額"]) if r["出入金"] == "入金" else -_f(r["交易金額"]), axis=1
    )
    df["key"] = df.apply(
        lambda r: f"TF|{r['date']}|{r['申請時間']}|{r['交易金額']}|{r['出入金']}", axis=1
    )
    return df[["key", "date", "signed"]].to_dict("records")


def parse_of_transactions():
    df = pd.read_html(OF_XLS_PATH, encoding="utf-8")[0]
    df["date"] = df["日期/時間"].str.split(" ").str[0].str.replace("/", "-")

    known_withdraw = {"出金", "預約出金"}
    known_skip = {"換匯"}
    unknown = set(df["移轉方式"]) - known_withdraw - known_skip - {"國內→國外 國外入金"}
    if unknown:
        print(f"[!] 海外出入金出現未知的移轉方式，這幾筆先當存款處理，請人工確認: {unknown}")

    def sign(row):
        method = row["移轉方式"]
        if method in known_skip:
            return 0.0
        if method in known_withdraw:
            return -_f(row["原幣金額"])
        return _f(row["原幣金額"])  # 國內→國外 國外入金、以及未知類型一律當存款

    df["signed"] = df.apply(sign, axis=1)
    df = df[df["signed"] != 0.0]
    df["key"] = df.apply(
        lambda r: f"OF|{r['日期/時間']}|{r['原幣金額']}|{r['移轉方式']}", axis=1
    )
    return df[["key", "date", "signed"]].to_dict("records")


def load_dw_log():
    if not os.path.exists(DW_LOG_PATH):
        return {}
    with open(DW_LOG_PATH, "r", encoding="utf-8") as f:
        records = json.load(f)
    return {r["key"]: r for r in records}


def merge_dw_log(existing, fresh_records):
    added = 0
    for r in fresh_records:
        if r["key"] not in existing:
            added += 1
        existing[r["key"]] = r
    return added


def save_dw_log(existing):
    records = sorted(existing.values(), key=lambda r: (r["date"], r["key"]))
    with open(DW_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    return records


def build_initial_and_map(records, market, history_start_date):
    market_records = [r for r in records if r["key"].startswith(market + "|")]
    by_date = {}
    for r in market_records:
        by_date[r["date"]] = by_date.get(r["date"], 0.0) + r["signed"]

    before = {d: v for d, v in by_date.items() if d < history_start_date}
    on_or_after = {d: v for d, v in by_date.items() if d >= history_start_date}

    if before:
        dates = ", ".join(f"{d}({v:+.0f})" for d, v in sorted(before.items()))
        print(f"    早於 history 起始日({history_start_date})的出入金併入起始權益: {dates}")

    return sum(before.values()), on_or_after


def backfill_market(history, by_date, initial_equity, equity_key, total_pl_key, dw_json_key, dw_map):
    all_dates = sorted(set(by_date.keys()) | set(dw_map.keys()))
    carry = initial_equity
    daily_pl = {}

    for d in all_dates:
        dw = dw_map.get(d, 0.0)
        rec = by_date.get(d)
        if rec is not None:
            rec[dw_json_key] = dw

        equity = _f(rec[equity_key]) if rec else 0.0
        total_pl = _f(rec.get(total_pl_key, 0.0)) if rec else 0.0

        if equity != 0:
            daily_pl[d] = equity - carry - dw
            carry = equity
        else:
            # equity 缺值：合成權益照常往前推進(供之後有真實equity的日子當基準)，
            # 但當天顯示的 daily_pl 退回用 total_pl(close_pl+float_pl)
            carry = carry + total_pl + dw
            if rec is not None:
                daily_pl[d] = total_pl

    return daily_pl


def main():
    with open(HISTORY_PATH, "r", encoding="utf-8") as f:
        history = json.load(f)

    history.sort(key=lambda r: r["date"])
    by_date = {r["date"]: r for r in history}
    history_start_date = history[0]["date"]

    dw_log = load_dw_log()
    print(f"[讀取] 出入金持久化紀錄: {len(dw_log)} 筆(來自之前累積的 {os.path.basename(DW_LOG_PATH)})")

    print("[讀取] 國內期貨出入金.xls")
    added_tf = merge_dw_log(dw_log, parse_tf_transactions())
    print(f"    新併入 {added_tf} 筆")

    print("[讀取] 海外期貨出入金.xls")
    added_of = merge_dw_log(dw_log, parse_of_transactions())
    print(f"    新併入 {added_of} 筆")

    records = save_dw_log(dw_log)
    print(f"[寫入] 出入金持久化紀錄共 {len(records)} 筆: {DW_LOG_PATH}")

    tf_base, tf_dw_map = build_initial_and_map(records, "TF", history_start_date)
    initial_tf_equity = BASE_EQUITY_BEFORE_RECORDS["TF"] + tf_base

    of_base, of_dw_map = build_initial_and_map(records, "OF", history_start_date)
    initial_of_equity = BASE_EQUITY_BEFORE_RECORDS["OF"] + of_base

    print(f"    INITIAL_TF_EQUITY = {initial_tf_equity:.0f}")
    print(f"    INITIAL_OF_EQUITY = {initial_of_equity:.0f}")

    tf_daily_pl = backfill_market(history, by_date, initial_tf_equity, "tf_equity", "tf_total_pl",
                                   "tf_deposit_withdrawal", tf_dw_map)
    of_daily_pl = backfill_market(history, by_date, initial_of_equity, "of_equity", "of_total_pl",
                                   "of_deposit_withdrawal", of_dw_map)

    for rec in history:
        rec.setdefault("tf_deposit_withdrawal", 0.0)
        rec.setdefault("of_deposit_withdrawal", 0.0)
        rec["tf_daily_pl_equity"] = tf_daily_pl.get(rec["date"])
        rec["of_daily_pl_equity"] = of_daily_pl.get(rec["date"])

    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    print(f"[完成] 已回填 {len(history)} 筆紀錄: {HISTORY_PATH}")


if __name__ == "__main__":
    main()
