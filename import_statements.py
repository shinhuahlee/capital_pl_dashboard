#%%
"""
對帳單匯入工具
將 F0200006398037_期貨沖銷明細.xlsx（國內）
與 F0200006398037_海外期貨沖銷明細_完整.xlsx（海外）
依 淨損益(參考) / 淨損益(台幣) 匯整每日損益，寫入 docs/history.json

執行方式：
    python import_statements.py
"""
import os
import json
import pandas as pd

REPO_DIR  = os.path.dirname(os.path.abspath(__file__))
DOM_FILE  = os.path.join(REPO_DIR, "F0200006398037_期貨沖銷明細.xlsx")
OVS_FILE  = os.path.join(REPO_DIR, "F0200006398037_海外期貨沖銷明細_完整.xlsx")
HIST_PATH = os.path.join(REPO_DIR, "docs", "history.json")

def to_float(v):
    if pd.isna(v):
        return 0.0
    return float(str(v).replace(",", "").strip())

def to_date_str(v):
    """YYYYMMDD (int or str) → '2026-05-28'"""
    s = str(int(float(str(v).replace(",", ""))))
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}"

def is_valid_date(v):
    try:
        n = int(float(str(v).replace(",", "")))
        return 20000101 <= n <= 20991231
    except Exception:
        return False

# ═══════════════════════════════════════════
# 1. 國內期貨沖銷明細
#    header: 沖銷日期, ..., 淨損益(參考)
# ═══════════════════════════════════════════
dom = pd.read_excel(DOM_FILE, header=0, engine="openpyxl", dtype=str)

# 去掉合計行、空行
date_col = "沖銷日期"
pl_col   = "淨損益(參考)"
dom = dom[dom[date_col].apply(is_valid_date)].copy()

dom["_date"] = dom[date_col].apply(to_date_str)
dom["_pl"]   = dom[pl_col].apply(to_float)

dom_daily = (
    dom.groupby("_date")["_pl"]
    .sum().reset_index()
    .rename(columns={"_date": "date", "_pl": "tf_close_pl"})
)

print(f"[國內] {len(dom)} 筆 → {len(dom_daily)} 個交易日")
for _, r in dom_daily.iterrows():
    print(f"  {r['date']}  淨損益(參考): {r['tf_close_pl']:>10,.0f}")

# ═══════════════════════════════════════════
# 2. 海外期貨沖銷明細
#    header: 平倉日期, ..., 淨損益(台幣)
# ═══════════════════════════════════════════
ovs = pd.read_excel(OVS_FILE, header=0, engine="openpyxl", dtype=str)

date_col2 = "平倉日期"
pl_col2   = "淨損益(台幣)"
ovs = ovs[ovs[date_col2].apply(is_valid_date)].copy()

ovs["_date"] = ovs[date_col2].apply(to_date_str)
ovs["_pl"]   = ovs[pl_col2].apply(to_float)

ovs_daily = (
    ovs.groupby("_date")["_pl"]
    .sum().reset_index()
    .rename(columns={"_date": "date", "_pl": "of_close_pl_twd"})
)

print(f"\n[海外] {len(ovs)} 筆 → {len(ovs_daily)} 個交易日")
for _, r in ovs_daily.iterrows():
    print(f"  {r['date']}  淨損益(台幣): {r['of_close_pl_twd']:>12,.0f}")

# ═══════════════════════════════════════════
# 3. 合併 → history record
# ═══════════════════════════════════════════
merged = pd.merge(dom_daily, ovs_daily, on="date", how="outer").fillna(0)
merged = merged.sort_values("date").reset_index(drop=True)

new_records = []
for _, row in merged.iterrows():
    tf = round(row["tf_close_pl"], 2)
    of = round(row["of_close_pl_twd"], 2)
    new_records.append({
        "date":        row["date"],
        "tf_close_pl": tf,
        "tf_float_pl": 0.0,
        "tf_total_pl": tf,
        "of_close_pl": of,   # 已換算台幣，ref_rate=1
        "of_float_pl": 0.0,
        "of_total_pl": of,
        "tf_equity":   0.0,
        "of_equity":   0.0,
        "of_currency": "NTD",
        "of_ref_rate": 1.0,
    })

# ═══════════════════════════════════════════
# 4. 與現有 history.json 合併
#    已有 API 資料的日期優先保留（含浮動損益）
# ═══════════════════════════════════════════
existing = []
if os.path.exists(HIST_PATH):
    with open(HIST_PATH, "r", encoding="utf-8") as f:
        existing = json.load(f)

existing_dates = {r["date"] for r in existing}

added, skipped = 0, 0
for rec in new_records:
    if rec["date"] not in existing_dates:
        existing.append(rec)
        added += 1
    else:
        skipped += 1

existing = sorted(existing, key=lambda x: x["date"])

with open(HIST_PATH, "w", encoding="utf-8") as f:
    json.dump(existing, f, ensure_ascii=False, indent=2)

print(f"\n[完成] 新增 {added} 天，跳過 {skipped} 天（已有API資料）")
print(f"  history.json 共 {len(existing)} 天")
print(f"  時間範圍：{existing[0]['date']} → {existing[-1]['date']}")

# 摘要
print("\n=== 每日損益摘要 ===")
cum = 0.0
for r in existing:
    tf = r["tf_total_pl"]
    of = r["of_total_pl"] * r["of_ref_rate"]
    total = tf + of
    cum += total
    print(f"  {r['date']}  國內:{tf:>10,.0f}  海外(TWD):{of:>11,.0f}  合計:{total:>10,.0f}  累積:{cum:>12,.0f}")

#%%