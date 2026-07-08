<?php
$dir = __DIR__ . '/wp-content/uploads/2026/06/';

$results = [];
for ($i = 1; $i <= 45; $i++) {
    $old = $dir . $i . '.jpg';
    $new = $dir . ($i + 400) . '.jpg';

    if (!file_exists($old)) {
        $results[] = "スキップ: {$i}.jpg が見つかりません";
        continue;
    }
    if (file_exists($new)) {
        $results[] = "スキップ: " . ($i + 400) . ".jpg はすでに存在します";
        continue;
    }
    if (rename($old, $new)) {
        $results[] = "OK: {$i}.jpg → " . ($i + 400) . ".jpg";
    } else {
        $results[] = "失敗: {$i}.jpg → " . ($i + 400) . ".jpg";
    }
}

echo "<pre>\n";
echo implode("\n", $results);
echo "\n</pre>";
