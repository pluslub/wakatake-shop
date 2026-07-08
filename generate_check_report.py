import csv
import html
import webbrowser
from pathlib import Path

SOURCE_FILES = [
    Path("商品データ/20260630summer_noimg.csv"),
    Path("商品データ/wakatake_new.csv"),
]
WOO_FILE = Path("商品データ/woo_export.csv")
OUT_PATH = Path("商品データ/check_report.html")

CHECK_COLS = ["名前", "標準価格", "カテゴリー", "画像"]


def load_csv(path):
    with open(path, encoding="utf-8-sig") as f:
        return {r["SKU"]: r for r in csv.DictReader(f) if r.get("SKU")}


def img_cell(url):
    if url:
        return (
            f'<img src="{url}" '
            f'onerror="this.outerHTML=\'<div class=no-img>404</div>\'">'
        )
    return "<div class='no-img'>URL<br>未設定</div>"


source_rows = {}
for f in SOURCE_FILES:
    source_rows.update(load_csv(f))

woo_rows = load_csv(WOO_FILE)

all_skus = sorted(set(source_rows) | set(woo_rows))

CSS = """
body{font-family:sans-serif;font-size:12px;margin:20px}
h1{font-size:18px}
h2{background:#2c3e50;color:white;padding:6px 12px;margin:28px 0 4px;border-radius:4px;font-size:14px}
table{border-collapse:collapse;width:100%;margin-bottom:8px}
th{background:#ecf0f1;padding:5px 8px;border:1px solid #ccc;text-align:left;white-space:nowrap}
td{padding:5px 8px;border:1px solid #ddd;vertical-align:middle}
tr.diff{background:#fff3cd}
tr.only-src{background:#d4edda}
tr.only-woo{background:#f8d7da}
.sku{font-weight:bold;white-space:nowrap}
.price{text-align:right;white-space:nowrap}
img{width:80px;height:80px;object-fit:contain;border:1px solid #ddd}
.no-img{width:80px;height:80px;border:2px solid red;color:red;font-size:10px;
        display:inline-flex;align-items:center;justify-content:center;text-align:center}
.ok{color:green;font-weight:bold}
.ng{color:red;font-weight:bold}
.tag{font-size:10px;padding:1px 5px;border-radius:3px;color:white;white-space:nowrap}
.tag-diff{background:#856404}
.tag-src{background:#155724}
.tag-woo{background:#721c24}
.legend{margin:8px 0;font-size:12px}
.legend span{display:inline-block;padding:3px 10px;margin-right:8px;border-radius:3px}
"""

lines = []
lines.append('<!DOCTYPE html>')
lines.append('<html lang="ja"><head><meta charset="UTF-8"><title>商品確認レポート</title>')
lines.append(f'<style>{CSS}</style></head><body>')
lines.append('<h1>商品データ確認レポート</h1>')
lines.append('<div class="legend">')
lines.append('<span style="background:#fff3cd;border:1px solid #ccc">差異あり</span>')
lines.append('<span style="background:#d4edda;border:1px solid #ccc">ソースCSVのみ（未インポート）</span>')
lines.append('<span style="background:#f8d7da;border:1px solid #ccc">WooCommerceのみ（ソースにない）</span>')
lines.append('</div>')

# --- 比較テーブル ---
lines.append('<h2>比較レポート（ソースCSV vs WooCommerce）</h2>')
lines.append('<table>')
lines.append('<tr><th>SKU</th><th>項目</th>'
             '<th>ソースCSV</th><th>WooCommerce</th><th>画像(ソース)</th><th>画像(WooCommerce)</th></tr>')

diff_count = 0
for sku in all_skus:
    src = source_rows.get(sku)
    woo = woo_rows.get(sku)

    if src and not woo:
        row_class = "only-src"
        tag = '<span class="tag tag-src">未インポート</span>'
    elif woo and not src:
        row_class = "only-woo"
        tag = '<span class="tag tag-woo">ソースなし</span>'
    else:
        diffs = [c for c in CHECK_COLS if c != "画像" and src.get(c, "").strip() != woo.get(c, "").strip()]
        row_class = "diff" if diffs else ""
        tag = '<span class="tag tag-diff">差異</span>' if diffs else '<span class="ok">✓</span>'
        if diffs:
            diff_count += 1

    src_name  = html.escape(src.get("名前", "") if src else "—")
    woo_name  = html.escape(woo.get("名前", "") if woo else "—")
    src_price = src.get("標準価格", "").strip() if src else ""
    woo_price = woo.get("標準価格", "").strip() if woo else ""
    src_img   = src.get("画像", "").strip() if src else ""
    woo_img   = woo.get("画像", "").strip() if woo else ""

    name_ok  = "ok" if src_name == woo_name else "ng"
    price_ok = "ok" if src_price == woo_price else "ng"

    src_price_disp = f"&yen;{int(src_price):,}" if src_price else "—"
    woo_price_disp = f"&yen;{int(woo_price):,}" if woo_price else "—"

    lines.append(f'<tr class="{row_class}">')
    lines.append(f'<td class="sku">{html.escape(sku)}<br>{tag}</td>')
    lines.append(f'<td><b class="{name_ok}">商品名</b><br><b class="{price_ok}">金額</b></td>')
    lines.append(f'<td>{src_name}<br><span class="price">{src_price_disp}</span></td>')
    lines.append(f'<td>{woo_name}<br><span class="price">{woo_price_disp}</span></td>')
    lines.append(f'<td>{img_cell(src_img)}</td>')
    lines.append(f'<td>{img_cell(woo_img)}</td>')
    lines.append('</tr>')

lines.append('</table>')
lines.append(f'<p>差異あり: <b>{diff_count}件</b> / 全{len(all_skus)}件</p>')

# --- ソースCSV 全商品一覧 ---
lines.append('<h2>ソースCSV 全商品一覧（目視確認用）</h2>')
lines.append('<table><tr><th>SKU</th><th>商品名</th><th>カテゴリー</th><th>画像</th><th>標準価格</th></tr>')
for sku in sorted(source_rows):
    r = source_rows[sku]
    price_raw = r.get("標準価格", "").strip()
    price_disp = f"&yen;{int(price_raw):,}" if price_raw else "—"
    lines.append(
        f'<tr><td class="sku">{html.escape(sku)}</td>'
        f'<td>{html.escape(r.get("名前",""))}</td>'
        f'<td>{html.escape(r.get("カテゴリー",""))}</td>'
        f'<td>{img_cell(r.get("画像","").strip())}</td>'
        f'<td class="price">{price_disp}</td></tr>'
    )
lines.append('</table>')

lines.append('</body></html>')

OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
print(f"Done: {OUT_PATH}")
webbrowser.open(str(OUT_PATH.resolve()))
