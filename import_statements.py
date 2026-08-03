"""
對帳單匯入工具
把國內/海外「沖銷明細」xlsx（用戶說這是最正確的損益數字來源）併入 docs/history.json。

國內: F0200006398037_期貨沖銷明細*.xlsx  (欄位: 沖銷日期, ..., 淨損益(參考))
海外: F0200006398037_海外期貨沖銷明細_完整*.xlsx (欄位: 平倉日期, ..., 淨損益(台幣))

這些 xlsx 每次從群益重新匯出都只涵蓋「當次匯出當下的一段區間」，較舊/較新的區間可能不會同時出現
在同一個檔案裡(跟出入金 xls 一樣的狀況)。所以資料夾裡會累積多個版本(無編號、2、3...)，
這支程式會把資料夾裡所有符合檔名規則的版本都讀進來，逐筆用唯一 key 去重後併入持久化的
settlement_log.json，之後 tf_close_pl/of_close_pl 一律從這份持久化 log 重新聚合。

跟舊版最大差異：舊版「日期已存在就跳過」，新版是「沖銷明細有涵蓋到的日期，一律覆蓋
tf_close_pl/of_close_pl(以及連帶的 tf_total_pl/of_total_pl)」，因為這是比較準確的來源；
tf_float_pl/of_float_pl(浮動損益)不是沖銷明細會有的東西，維持原本 history.json 裡的值不動
(如果該日期本來就沒有紀錄，才會建立新的一列、float_pl 設為 0)。

執行方式：
    python import_statements.py
"""
import glob
import json
import os

import pandas as pd

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
HIST_PATH = os.path.join(REPO_DIR, "docs", "history.json")
SETTLEMENT_LOG_PATH = os.path.join(REPO_DIR, "settlement_log.json")

DOM_GLOB = os.path.join(REPO_DIR, "F0200006398037_期貨沖銷明細*.xlsx")
# 海外檔名裡也含「期貨沖銷明細」，用「海外」開頭排除跟國內混到
OVS_GLOB = os.path.join(REPO_DIR, "F0200006398037_海外期貨沖銷明細_完整*.xlsx")


def to_float(v):
    if pd.isna(v):
        return 0.0
    return float(str(v).replace(",", "").strip())


def is_valid_date(v):
    try:
        n = int(float(str(v).replace(",", "")))
        return 20000101 <= n <= 20991231
    except Exception:
        return False


def to_date_str(v):
    s = str(int(float(str(v).replace(",", ""))))
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}"


def parse_dom_file(path):
    df = pd.read_excel(path, header=0, engine="openpyxl", dtype=str)
    df = df[df["沖銷日期"].apply(is_valid_date)].copy()
    df["date"] = df["沖銷日期"].apply(to_date_str)
    df["signed"] = df["淨損益(參考)"].apply(to_float)
    df["key"] = df.apply(
        lambda r: "TF|" + "|".join(str(r[c]) for c in
                                    ["沖銷日期", "成交日", "成交時間", "商品", "買賣", "成交價", "成交量", "淨損益(參考)"]),
        axis=1,
    )
    return df[["key", "date", "signed"]].to_dict("records")


def parse_ovs_file(path):
    df = pd.read_excel(path, header=0, engine="openpyxl", dtype=str)
    df = df[df["平倉日期"].apply(is_valid_date)].copy()
    df["date"] = df["平倉日期"].apply(to_date_str)
    df["signed"] = df["淨損益(台幣)"].apply(to_float)
    df["key"] = df.apply(
        lambda r: "OF|" + "|".join(str(r[c]) for c in
                                    ["平倉日期", "成交日期", "商品名稱", "商品年月", "買賣別", "口數", "成交價", "淨損益(台幣)"]),
        axis=1,
    )
    return df[["key", "date", "signed"]].to_dict("records")


def load_log():
    if not os.path.exists(SETTLEMENT_LOG_PATH):
        return {}
    with open(SETTLEMENT_LOG_PATH, "r", encoding="utf-8") as f:
        return {r["key"]: r for r in json.load(f)}


def merge_log(existing, fresh_records):
    added = 0
    for r in fresh_records:
        if r["key"] not in existing:
            added += 1
        existing[r["key"]] = r
    return added


def save_log(existing):
    records = sorted(existing.values(), key=lambda r: (r["date"], r["key"]))
    with open(SETTLEMENT_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    return records


def main():
    log = load_log()
    print(f"[讀取] 持久化沖銷紀錄: {len(log)} 筆")

    dom_files = sorted(glob.glob(DOM_GLOB))
    dom_files = [p for p in dom_files if "海外" not in os.path.basename(p)]
    ovs_files = sorted(glob.glob(OVS_GLOB))

    for path in dom_files:
        added = merge_log(log, parse_dom_file(path))
        print(f"[國內] {os.path.basename(path)}  新併入 {added} 筆")

    for path in ovs_files:
        added = merge_log(log, parse_ovs_file(path))
        print(f"[海外] {os.path.basename(path)}  新併入 {added} 筆")

    records = save_log(log)
    print(f"[寫入] 持久化沖銷紀錄共 {len(records)} 筆: {SETTLEMENT_LOG_PATH}")

    tf_daily = {}
    of_daily = {}
    for r in records:
        d, market, signed = r["date"], r["key"].split("|", 1)[0], r["signed"]
        bucket = tf_daily if market == "TF" else of_daily
        bucket[d] = bucket.get(d, 0.0) + signed

    with open(HIST_PATH, "r", encoding="utf-8") as f:
        history = json.load(f)
    by_date = {r["date"]: r for r in history}

    updated, created = 0, 0
    for d, tf_close in tf_daily.items():
        rec = by_date.get(d)
        if rec is None:
            rec = {"date": d, "tf_float_pl": 0.0, "of_close_pl": 0.0, "of_float_pl": 0.0,
                   "tf_equity": 0.0, "of_equity": 0.0, "of_currency": "NTD", "of_ref_rate": 1.0}
            history.append(rec)
            by_date[d] = rec
            created += 1
        else:
            updated += 1
        rec["tf_close_pl"] = round(tf_close, 2)
        rec["tf_total_pl"] = round(tf_close + rec.get("tf_float_pl", 0.0), 2)

    for d, of_close in of_daily.items():
        rec = by_date.get(d)
        if rec is None:
            rec = {"date": d, "tf_close_pl": 0.0, "tf_float_pl": 0.0, "of_float_pl": 0.0,
                   "tf_equity": 0.0, "of_equity": 0.0, "of_currency": "NTD", "of_ref_rate": 1.0}
            history.append(rec)
            by_date[d] = rec
            created += 1
        else:
            updated += 1
        rec["of_close_pl"] = round(of_close, 2)
        rec["of_total_pl"] = round(of_close + rec.get("of_float_pl", 0.0), 2)

    history = sorted(history, key=lambda r: r["date"])
    with open(HIST_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    print(f"[完成] history.json 覆蓋/新增 close_pl 共 {updated + created} 個日期記錄"
          f"(更新 {updated}，新增 {created})，總計 {len(history)} 天")
    print("[提醒] tf_close_pl/of_close_pl 已更新，記得重新執行 backfill_daily_pl.py 讓權益差額法用新數字重算")


if __name__ == "__main__":
    main()
