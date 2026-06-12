"""
テスト注文削除スクリプト
test_order_ids.json に記録された注文をすべて完全削除します。
"""
import requests
import json
from requests.auth import HTTPBasicAuth

WC_URL    = "https://wakatake.info/shop"
WC_KEY    = "ck_bb753c4e1e11e7b08fc3050adca49cdf16e4e393"
WC_SECRET = "cs_b33472c68e5f3dc6be8cddb7cd1a93d1c63ae789"

auth = HTTPBasicAuth(WC_KEY, WC_SECRET)
BASE = f"{WC_URL}/wp-json/wc/v3"


def main():
    try:
        with open("test_order_ids.json") as f:
            ids = json.load(f)
    except FileNotFoundError:
        print("test_order_ids.json が見つかりません。")
        return

    print("=" * 50)
    print("  ⚠️  警告：本番環境への操作です")
    print(f"  接続先: {WC_URL}")
    print(f"  削除件数: {len(ids)}件（完全削除・復元不可）")
    print("=" * 50)
    ans = input("\n本当に削除しますか？ (yes と入力して Enter): ").strip()
    if ans != "yes":
        print("キャンセルしました。")
        return
    print()
    print(f"{len(ids)}件の注文を削除します...")
    deleted = 0
    for oid in ids:
        r = requests.delete(f"{BASE}/orders/{oid}", auth=auth, params={"force": True})
        if r.status_code == 200:
            print(f"  削除: #{oid}")
            deleted += 1
        else:
            print(f"  失敗: #{oid} → {r.status_code}")

    print(f"\n完了: {deleted}/{len(ids)}件削除")


if __name__ == "__main__":
    main()
