"""
WooCommerce テスト注文生成スクリプト
作成した注文IDは test_order_ids.json に保存されます。
削除は delete_test_orders.py で行えます。
"""
import requests
import random
import json
from requests.auth import HTTPBasicAuth

WC_URL    = "https://wakatake.info/shop"
WC_KEY    = "ck_bb753c4e1e11e7b08fc3050adca49cdf16e4e393"
WC_SECRET = "cs_b33472c68e5f3dc6be8cddb7cd1a93d1c63ae789"

auth = HTTPBasicAuth(WC_KEY, WC_SECRET)
BASE = f"{WC_URL}/wp-json/wc/v3"

DEPTS = ["総務部", "営業部", "製造部", "経理部", "人事部", "技術部", "企画部", "広報部"]
NAMES = [
    ("田中", "太郎"), ("佐藤", "花子"), ("山田", "次郎"), ("鈴木", "美咲"),
    ("高橋", "一郎"), ("伊藤", "由美"), ("渡辺", "健二"), ("中村", "さくら"),
    ("小林", "浩二"), ("加藤", "恵子"), ("吉田", "誠"),  ("山本", "律子"),
]

ORDER_COUNT = 50
PRODUCTS_PER_ORDER = (2, 6)   # 1注文あたりの商品種類数（min, max）
QTY_RANGE = (1, 15)            # 1商品あたりの数量


def get_all_products():
    products = []
    page = 1
    while True:
        r = requests.get(
            f"{BASE}/products",
            auth=auth,
            params={"per_page": 100, "page": page, "status": "publish"},
        )
        r.raise_for_status()
        data = r.json()
        if not data:
            break
        products.extend(data)
        print(f"  商品取得: {len(products)}件...")
        if len(data) < 100:
            break
        page += 1
    return products


def build_order(product_ids):
    last, first = random.choice(NAMES)
    dept = random.choice(DEPTS)
    selected = random.sample(product_ids, min(random.randint(*PRODUCTS_PER_ORDER), len(product_ids)))
    line_items = [{"product_id": pid, "quantity": random.randint(*QTY_RANGE)} for pid in selected]
    return {
        "payment_method":       "cheque",
        "payment_method_title": "振込",
        "set_paid":             True,
        "status":               "completed",
        "billing": {
            "first_name": first,
            "last_name":  last,
            "company":    dept,
            "address_1":  "1-1-1",
            "city":       "大阪市",
            "state":      "JP27",
            "postcode":   "530-0001",
            "country":    "JP",
            "email":      "test-order@example.com",
            "phone":      "06-0000-0000",
        },
        "line_items": line_items,
    }


def main():
    print("=" * 50)
    print("  ⚠️  警告：本番環境への操作です")
    print(f"  接続先: {WC_URL}")
    print(f"  作成件数: {ORDER_COUNT}件")
    print()
    print("  【メール通知について】")
    print("  ✅ 顧客へのメール : 送信されません")
    print("       （注文者メールは test-order@example.com（架空）を使用）")
    print("  ⚠️  管理者へのメール: 注文通知が届く可能性があります")
    print("       （WooCommerce → 設定 → メール → 新しい注文 で一時停止可能）")
    print()
    print("  ※ 作成した注文は delete_test_orders.py で削除できます")
    print("=" * 50)
    ans = input("\n本当に実行しますか？ (yes と入力して Enter): ").strip()
    if ans != "yes":
        print("キャンセルしました。")
        return
    print()
    print("商品一覧を取得中...")
    products = get_all_products()
    purchasable = [p["id"] for p in products if p.get("purchasable") and p.get("stock_status") != "outofstock"]
    print(f"購入可能な商品: {len(purchasable)}件\n")

    if not purchasable:
        print("商品が見つかりません。中止します。")
        return

    created_ids = []
    for i in range(ORDER_COUNT):
        order_data = build_order(purchasable)
        r = requests.post(f"{BASE}/orders", auth=auth, json=order_data)
        if r.status_code == 201:
            oid = r.json()["id"]
            created_ids.append(oid)
            print(f"  [{i+1:02d}/{ORDER_COUNT}] 注文 #{oid} 作成")
        else:
            print(f"  [{i+1:02d}/{ORDER_COUNT}] 失敗: {r.status_code} {r.text[:100]}")

    with open("test_order_ids.json", "w") as f:
        json.dump(created_ids, f)

    print(f"\n完了: {len(created_ids)}件作成 → test_order_ids.json に保存")


if __name__ == "__main__":
    main()
