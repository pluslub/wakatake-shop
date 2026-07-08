Add-Type -AssemblyName System.Web

$csvFiles = @(
    ".\商品データ\20260630summer_noimg.csv",
    ".\商品データ\wakatake_new.csv"
)

$rows = @()
foreach ($file in $csvFiles) {
    $rows += Import-Csv $file -Encoding UTF8
}

$rows = $rows | Sort-Object { $_.'カテゴリー' }, { $_.'SKU' }

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('<!DOCTYPE html>')
$lines.Add('<html lang="ja"><head><meta charset="UTF-8"><title>商品確認レポート</title>')
$lines.Add('<style>')
$lines.Add('body{font-family:sans-serif;font-size:13px;margin:20px}')
$lines.Add('h2{background:#2c3e50;color:white;padding:8px 12px;margin:24px 0 4px;border-radius:4px}')
$lines.Add('table{border-collapse:collapse;width:100%}')
$lines.Add('th{background:#ecf0f1;padding:6px 10px;border:1px solid #ccc;text-align:left}')
$lines.Add('td{padding:6px 10px;border:1px solid #ddd;vertical-align:middle}')
$lines.Add('tr:nth-child(even){background:#f9f9f9}')
$lines.Add('.sku{font-weight:bold;white-space:nowrap}')
$lines.Add('.price{text-align:right;white-space:nowrap}')
$lines.Add('img{width:100px;height:100px;object-fit:contain;border:1px solid #ddd}')
$lines.Add('.no-img{width:100px;height:100px;border:2px solid red;color:red;font-size:11px;display:flex;align-items:center;justify-content:center;text-align:center}')
$lines.Add('</style></head><body>')
$lines.Add('<h1>商品データ確認レポート</h1>')
$lines.Add('<p>確認項目: SKU / 商品名 / 画像 / 金額</p>')

$currentCategory = $null

foreach ($row in $rows) {
    $category = $row.'カテゴリー'

    if ($category -ne $currentCategory) {
        if ($currentCategory -ne $null) {
            $lines.Add('</table>')
        }
        $label = if ($category) { $category } else { '（カテゴリーなし）' }
        $lines.Add("<h2>$label</h2>")
        $lines.Add('<table><tr><th>SKU</th><th>商品名</th><th>画像</th><th>標準価格</th></tr>')
        $currentCategory = $category
    }

    $sku   = [System.Web.HttpUtility]::HtmlEncode($row.'SKU')
    $name  = [System.Web.HttpUtility]::HtmlEncode($row.'名前')
    $price = $row.'標準価格'
    $imgUrl = $row.'画像'

    $priceDisplay = if ($price) { '&yen;' + ('{0:N0}' -f [int]$price) } else { '-' }

    if ($imgUrl) {
        $imgCell = "<img src='$imgUrl' onerror=" + '"' + "this.outerHTML='<div class=no-img>404</div>'" + '"' + '>'
    } else {
        $imgCell = "<div class='no-img'>URL未設定</div>"
    }

    $lines.Add("<tr><td class='sku'>$sku</td><td>$name</td><td>$imgCell</td><td class='price'>$priceDisplay</td></tr>")
}

if ($currentCategory -ne $null) {
    $lines.Add('</table>')
}
$lines.Add('</body></html>')

$outPath = ".\商品データ\check_report.html"
[System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.UTF8Encoding]::new($false))
Write-Host "生成完了: $outPath"
Start-Process $outPath
