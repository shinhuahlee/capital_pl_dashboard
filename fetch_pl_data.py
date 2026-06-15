"""
群益 API 損益資料抓取器
使用 SKDLLPythonTester/SKDLLPython.py 的 SK class (DLL 直接呼叫版)

使用方式:
  python fetch_pl_data.py --id 你的帳號 --pw 你的密碼

輸出:
  docs/data.json  (供 GitHub Pages HTML 讀取)

前置作業:
  1. 安裝群益策略王
  2. 將 python_winner/3.群益pythonAPI/SKDLLPythonTester 整個資料夾複製到本專案根目錄
     或修改下方 TESTER_DIR 路徑
"""

import sys
import os
import json
import time
import threading
import argparse
from datetime import datetime

# ── 路徑設定：指向 SKDLLPythonTester 資料夾 ──
TESTER_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "SKDLLPythonTester")

# 如果本資料夾沒有，自動嘗試 python_winner 裡的
if not os.path.exists(TESTER_DIR):
    TESTER_DIR = r"C:\GitHub\python_winner\3.群益pythonAPI\SKDLLPythonTester"

if not os.path.exists(TESTER_DIR):
    print("[!] 找不到 SKDLLPythonTester 資料夾，請確認路徑設定。")
    sys.exit(1)

sys.path.insert(0, TESTER_DIR)
_orig_cwd = os.getcwd()
os.chdir(TESTER_DIR)       # SK class 用 __file__ 相對路徑找 SKCOM.dll
from SKDLLPython import SK
os.chdir(_orig_cwd)

# ── 參數 ──
parser = argparse.ArgumentParser(description="群益 API 損益抓取")
parser.add_argument("--id",  required=True, help="群益登入 ID")
parser.add_argument("--pw",  required=True, help="群益密碼")
parser.add_argument("--out", default=os.path.join("docs", "data.json"),
                    help="輸出 JSON 路徑 (預設: docs/data.json)")
args = parser.parse_args()

LOGIN_ID = args.id
PASSWORD = args.pw
OUT_PATH = args.out

# ── 等待登入完成 ──
login_done  = threading.Event()
tf_accounts = []
of_accounts = []

def _on_connection(login_id, code):
    print(f"[OnConnection] {login_id}  {SK.GetMessage(code)}")
    if code == 0:
        login_done.set()

SK.OnConnection(_on_connection)

print(f"[*] 登入中: {LOGIN_ID} ...")
result = SK.Login(LOGIN_ID, PASSWORD)
print(f"[*] Login: {SK.GetMessage(result.Code)}")

if result.Code != 0:
    print("[!] 登入失敗。")
    sys.exit(1)

tf_accounts = result.TFAccounts
of_accounts = result.OFAccounts
print(f"[*] TF 帳號: {[a.FullAccount for a in tf_accounts]}")
print(f"[*] OF 帳號: {[a.FullAccount for a in of_accounts]}")

login_done.wait(timeout=30)

# ── bytes → str ──
def _s(val):
    if val is None:
        return ""
    if isinstance(val, bytes):
        return val.decode("ansi", errors="replace").strip()
    return str(val).strip()

# ── 抓取函式 ──

def fetch_tf_rights(login_id, full_account):
    res = SK.GetFutureRights(login_id, full_account, 0)
    if res.StatusCode != 0 or not res.Blocks:
        print(f"[!] GetFutureRights: {res.Message}")
        return None
    b = res.Blocks[0]
    return {
        "account":             full_account,
        "account_balance":     _s(b.AccountBalance),
        "floating_pl":         _s(b.FloatingPL),
        "futures_close_pl":    _s(b.FuturesClosePL),
        "intraday_unrealized": _s(b.IntradayUnrealized),
        "equity":              _s(b.Equity),
        "excess_margin":       _s(b.ExcessMargin),
        "initial_margin":      _s(b.InitialMargin),
        "maintenance_margin":  _s(b.MaintenanceMargin),
        "risk_indicator":      _s(b.RiskIndicator),
        "yesterday_balance":   _s(b.YesterdayBalance),
        "realized_fee":        _s(b.RealizedFee),
        "transaction_tax":     _s(b.TransactionTax),
        "available_balance":   _s(b.AvailableBalance),
    }


def fetch_of_rights(login_id, full_account):
    res = SK.GetOFFutureRights(login_id, full_account, 0)
    if res.StatusCode != 0 or not res.Blocks:
        print(f"[!] GetOFFutureRights: {res.Message}")
        return None
    b = res.Blocks[0]
    return {
        "account":               full_account,
        "currency":              _s(b.Currency),
        "account_equity":        _s(b.AccountEquity),
        "account_balance":       _s(b.AccountBalance),
        "futures_floating_pl":   _s(b.FuturesFloatingPL),
        "futures_close_pl_today":_s(b.FuturesClosePLToday),
        "intraday_unrealized":   _s(b.IntradayUnrealized),
        "initial_margin":        _s(b.InitialMargin),
        "maintenance_margin":    _s(b.MaintenanceMargin),
        "risk_indicator":        _s(b.RiskIndicator),
        "withdrawable_margin":   _s(b.WithdrawableMargin),
        "realized_fee":          _s(b.RealizedFee),
        "previous_day_balance":  _s(b.PreviousDayBalance),
        "reference_rate":        _s(b.ReferenceRate),
    }


def fetch_tf_positions(login_id, full_account):
    res = SK.GetOpenInterestGW(login_id, full_account, 0)
    if res.StatusCode != 0:
        print(f"[!] GetOpenInterestGW: {res.Message}")
        return []
    return [{
        "market_type": _s(b.MarketType),
        "account":     _s(b.Account),
        "future_no":   _s(b.FutureNo),
        "buy_sell":    _s(b.BuySell),
        "qty":         _s(b.Qty),
        "qty_trade":   _s(b.Qty_Trade),
        "avg_price":   _s(b.AvgPrice),
        "fee":         _s(b.Fee),
        "tax":         _s(b.Tax),
    } for b in res.Blocks]


def fetch_of_positions(login_id, full_account):
    res = SK.GetOFOpenInterestGW(login_id, full_account, 0)
    if res.StatusCode != 0:
        print(f"[!] GetOFOpenInterestGW: {res.Message}")
        return []
    return [{
        "account":         _s(b.Account),
        "exchange_code":   _s(b.ExchangeCode),
        "exchange_name":   _s(b.ExchangeName),
        "product_code_ym": _s(b.ProductCodeWithYM),
        "product_code":    _s(b.ProductCode),
        "product_name_ym": _s(b.ProductNameWithYM),
        "buy_sell":        _s(b.BuySell),
        "open_interest":   _s(b.OpenInterest),
        "avg_price":       _s(b.AvgPrice),
        "market_price":    _s(b.MarketPrice),
        "unrealized_pl":   _s(b.UnrealizedPL),
        "prev_settlement": _s(b.PrevSettlementPrice),
        "initial_margin":  _s(b.InitialMargin),
        "is_option":       _s(b.IsOption),
        "strike_price":    _s(b.StrikePrice),
        "call_put":        _s(b.CallPut),
    } for b in res.Blocks]


# ── 主程式 ──
output = {
    "updated_at":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "tf_rights":    [],
    "of_rights":    [],
    "tf_positions": [],
    "of_positions": [],
}

for acc in tf_accounts:
    print(f"[*] TF 權益: {acc.FullAccount}")
    r = fetch_tf_rights(acc.LoginID, acc.FullAccount)
    if r:
        output["tf_rights"].append(r)
    output["tf_positions"].extend(fetch_tf_positions(acc.LoginID, acc.FullAccount))

for acc in of_accounts:
    print(f"[*] OF 權益: {acc.FullAccount}")
    r = fetch_of_rights(acc.LoginID, acc.FullAccount)
    if r:
        output["of_rights"].append(r)
    output["of_positions"].extend(fetch_of_positions(acc.LoginID, acc.FullAccount))

os.makedirs(os.path.dirname(OUT_PATH) if os.path.dirname(OUT_PATH) else ".", exist_ok=True)
with open(OUT_PATH, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"\n[完成] 已寫入: {OUT_PATH}")
