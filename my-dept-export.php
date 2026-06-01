<?php
/**
 * Plugin Name: 部署別・カテゴリ別・CSVエクスポート（独立ボタン版）
 * Description: 部署別と商品カテゴリ別、それぞれの専用ボタンでCSVを書き出します。
 * Version: 1.7
 */

if (!defined('ABSPATH')) exit;

// メニュー登録
add_action('admin_menu', function() {
    add_menu_page('CSV出力', 'CSV出力', 'manage_woocommerce', 'my-excel-export', 'render_excel_export_page', 'dashicons-media-spreadsheet', 56);
});

// 操作画面
function render_excel_export_page() {
    global $wpdb;

    // 部署一覧の取得
    $depts = $wpdb->get_results("
        SELECT meta_value as dept, COUNT(post_id) as order_count
        FROM {$wpdb->postmeta}
        WHERE meta_key = '_billing_company' AND meta_value != ''
        GROUP BY meta_value ORDER BY meta_value ASC
    ");

    // カテゴリ別の注文数取得
    $cat_counts = $wpdb->get_results("
        SELECT t.name, t.slug, COUNT(DISTINCT p.ID) as order_count
        FROM {$wpdb->prefix}terms t
        INNER JOIN {$wpdb->prefix}term_taxonomy tt ON t.term_id = tt.term_id
        INNER JOIN {$wpdb->prefix}term_relationships tr ON tt.term_taxonomy_id = tr.term_taxonomy_id
        INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta woim ON tr.object_id = woim.meta_value
        INNER JOIN {$wpdb->prefix}woocommerce_order_items woi ON woim.order_item_id = woi.order_item_id
        INNER JOIN {$wpdb->posts} p ON woi.order_id = p.ID
        WHERE tt.taxonomy = 'product_cat'
          AND woim.meta_key = '_product_id'
          AND p.post_type = 'shop_order'
          AND p.post_status IN ('wc-processing', 'wc-completed')
        GROUP BY t.term_id ORDER BY t.name ASC
    ");
?>
    <div class="wrap">
        <h1>注文CSVエクスポート</h1>

        <div style="display: flex; flex-wrap: wrap; gap: 20px;">

            <!-- 部署別 -->
            <div style="flex: 1 1 calc(50% - 10px); padding: 20px; border: 1px solid #ccc; background: #fff;">
                <h2>部署別に出力</h2>
                <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
                    <input type="hidden" name="action" value="my_excel_csv_export">
                    <input type="hidden" name="export_type" value="dept">
                    <?php wp_nonce_field('my_excel_export_nonce'); ?>

                    <div style="margin-bottom: 6px;">
                        <label style="font-weight: bold;">
                            <input type="checkbox" id="dept-check-all" onchange="
                                document.querySelectorAll('.dept-item').forEach(cb => cb.checked = this.checked);
                            "> 全て
                        </label>
                    </div>
                    <div style="max-height: 220px; overflow-y: auto; border: 1px solid #ddd; padding: 8px; margin-bottom: 10px;">
                        <?php foreach ($depts as $row): ?>
                            <label style="display: block; margin-bottom: 4px;">
                                <input type="checkbox" class="dept-item" name="target_val[]" value="<?php echo esc_attr($row->dept); ?>" onchange="
                                    var all = document.querySelectorAll('.dept-item');
                                    var checked = document.querySelectorAll('.dept-item:checked');
                                    document.getElementById('dept-check-all').checked = all.length === checked.length;
                                ">
                                <?php echo esc_html($row->dept); ?> （<?php echo (int)$row->order_count; ?>件）
                            </label>
                        <?php endforeach; ?>
                    </div>
                    <?php submit_button('部署別CSVをダウンロード'); ?>
                </form>
            </div>

            <!-- カテゴリ別 -->
            <div style="flex: 1 1 calc(50% - 10px); padding: 20px; border: 1px solid #ccc; background: #fff;">
                <h2>商品カテゴリ別に出力</h2>
                <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
                    <input type="hidden" name="action" value="my_excel_csv_export">
                    <input type="hidden" name="export_type" value="category">
                    <?php wp_nonce_field('my_excel_export_nonce'); ?>

                    <div style="margin-bottom: 6px;">
                        <label style="font-weight: bold;">
                            <input type="checkbox" id="cat-check-all" onchange="
                                document.querySelectorAll('.cat-item').forEach(cb => cb.checked = this.checked);
                            "> 全て
                        </label>
                    </div>
                    <div style="max-height: 220px; overflow-y: auto; border: 1px solid #ddd; padding: 8px; margin-bottom: 10px;">
                        <?php foreach ($cat_counts as $cat): ?>
                            <label style="display: block; margin-bottom: 4px;">
                                <input type="checkbox" class="cat-item" name="target_val[]" value="<?php echo esc_attr($cat->slug); ?>" onchange="
                                    var all = document.querySelectorAll('.cat-item');
                                    var checked = document.querySelectorAll('.cat-item:checked');
                                    document.getElementById('cat-check-all').checked = all.length === checked.length;
                                ">
                                <?php echo esc_html($cat->name); ?> （<?php echo (int)$cat->order_count; ?>件）
                            </label>
                        <?php endforeach; ?>
                    </div>
                    <?php submit_button('カテゴリ別CSVをダウンロード', 'secondary'); ?>
                </form>
            </div>

            <!-- 納品書用 -->
            <div style="flex: 1 1 calc(50% - 10px); padding: 20px; border: 1px solid #ccc; background: #fff;">
                <h2>納品書用に出力</h2>
                <p>処理中・完了の全注文を商品ごとに集計して出力します。</p>
                <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
                    <input type="hidden" name="action" value="my_excel_csv_export">
                    <input type="hidden" name="export_type" value="delivery">
                    <?php wp_nonce_field('my_excel_export_nonce'); ?>
                    <?php submit_button('納品書用CSVをダウンロード', 'secondary'); ?>
                </form>
            </div>

            <!-- 氏名検索リンク -->
            <div style="flex: 1 1 calc(50% - 10px); padding: 20px; border: 1px solid #ccc; background: #fff;">
                <h2>氏名で検索</h2>
                <p>氏名で注文を調べる場合はWooCommerceの注文画面から検索できます。</p>
                <a href="<?php echo admin_url('edit.php?post_type=shop_order'); ?>" class="button button-secondary">氏名検索はこちら</a>
            </div>

        </div>
    </div>
<?php
}

// カテゴリのスラッグから注文IDを取得するヘルパー
function my_export_get_order_ids_by_category(array $slugs): array {
    global $wpdb;

    $statuses = ['wc-processing', 'wc-completed'];

    $slug_ph = implode(',', array_fill(0, count($slugs), '%s'));
    $term_ids = $wpdb->get_col($wpdb->prepare(
        "SELECT term_id FROM {$wpdb->terms} WHERE slug IN ($slug_ph)",
        ...$slugs
    ));
    if (empty($term_ids)) return [];

    $term_ph = implode(',', array_fill(0, count($term_ids), '%d'));
    $product_ids = $wpdb->get_col($wpdb->prepare(
        "SELECT DISTINCT object_id
         FROM {$wpdb->term_relationships} tr
         INNER JOIN {$wpdb->term_taxonomy} tt ON tr.term_taxonomy_id = tt.term_taxonomy_id
         WHERE tt.term_id IN ($term_ph) AND tt.taxonomy = 'product_cat'",
        ...$term_ids
    ));
    if (empty($product_ids)) return [];

    $prod_ph   = implode(',', array_fill(0, count($product_ids), '%d'));
    $status_ph = implode(',', array_fill(0, count($statuses), '%s'));
    $order_ids = $wpdb->get_col($wpdb->prepare(
        "SELECT DISTINCT woi.order_id
         FROM {$wpdb->prefix}woocommerce_order_items woi
         INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta woim ON woi.order_item_id = woim.order_item_id
         INNER JOIN {$wpdb->posts} p ON woi.order_id = p.ID
         WHERE woim.meta_key = '_product_id'
           AND woim.meta_value IN ($prod_ph)
           AND p.post_type = 'shop_order'
           AND p.post_status IN ($status_ph)",
        ...array_merge($product_ids, $statuses)
    ));

    return array_map('intval', $order_ids);
}

// 実行ロジック
add_action('admin_post_my_excel_csv_export', function() {
    global $wpdb;
    check_admin_referer('my_excel_export_nonce');

    $type = $_POST['export_type'];

    // 納品書用：商品ごとに集計して直接出力
    if ($type === 'delivery') {
        $rows = $wpdb->get_results("
            SELECT
                woi.order_item_name as product_name,
                SUM(CAST(woim_qty.meta_value AS UNSIGNED)) as total_qty,
                SUM(CAST(woim_total.meta_value AS DECIMAL(10,2)) + CAST(woim_tax.meta_value AS DECIMAL(10,2))) as total_amount
            FROM {$wpdb->prefix}woocommerce_order_items woi
            INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta woim_qty
                ON woi.order_item_id = woim_qty.order_item_id AND woim_qty.meta_key = '_qty'
            INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta woim_total
                ON woi.order_item_id = woim_total.order_item_id AND woim_total.meta_key = '_line_total'
            INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta woim_tax
                ON woi.order_item_id = woim_tax.order_item_id AND woim_tax.meta_key = '_line_tax'
            INNER JOIN {$wpdb->posts} o ON woi.order_id = o.ID
            WHERE woi.order_item_type = 'line_item'
              AND o.post_type = 'shop_order'
              AND o.post_status IN ('wc-processing', 'wc-completed')
            GROUP BY woi.order_item_name
            ORDER BY woi.order_item_name ASC
        ");

        header('Content-Type: text/csv; charset=UTF-8');
        header('Content-Disposition: attachment; filename="' . urlencode('納品書_' . date('Ymd')) . '.csv"');

        $output = fopen('php://output', 'w');
        fwrite($output, "\xEF\xBB\xBF");
        fputcsv($output, ['商品名', '単価(税込)', '合計数量', '合計金額(税込)']);
        foreach ($rows as $row) {
            $qty        = (int)$row->total_qty;
            $total      = (float)$row->total_amount;
            $unit_price = $qty > 0 ? round($total / $qty) : 0;
            fputcsv($output, [
                $row->product_name,
                number_format($unit_price, 0),
                $qty,
                number_format($total, 0),
            ]);
        }
        fclose($output);
        exit;
    }

    // 部署別・カテゴリ別：注文単位で出力
    $base_args = ['limit' => -1, 'status' => ['wc-processing', 'wc-completed']];
    $orders    = [];

    if ($type === 'category') {
        $vals = array_map('sanitize_text_field', (array)($_POST['target_val'] ?? []));
        if (empty($vals)) wp_die('カテゴリを1つ以上選択してください。');

        $order_ids = my_export_get_order_ids_by_category($vals);
        foreach ($order_ids as $id) {
            $order = wc_get_order($id);
            if ($order) $orders[] = $order;
        }
        $label    = count($vals) === 1 ? $vals[0] : $vals[0] . '_他' . (count($vals) - 1) . '件';
        $filename = "カテゴリ別_" . $label;

    } else { // dept
        $vals = array_map('sanitize_text_field', (array)($_POST['target_val'] ?? []));
        if (empty($vals)) wp_die('部署を1つ以上選択してください。');

        $all_orders = [];
        foreach ($vals as $v) {
            foreach (wc_get_orders(array_merge($base_args, ['billing_company' => $v])) as $order) {
                $all_orders[$order->get_id()] = $order;
            }
        }
        $orders   = array_values($all_orders);
        $label    = count($vals) === 1 ? $vals[0] : $vals[0] . '_他' . (count($vals) - 1) . '件';
        $filename = "部署別_" . $label;
    }

    header('Content-Type: text/csv; charset=UTF-8');
    header('Content-Disposition: attachment; filename="' . urlencode($filename) . '_' . date('Ymd') . '.csv"');

    $output = fopen('php://output', 'w');
    fwrite($output, "\xEF\xBB\xBF");
    fputcsv($output, ['注文番号', '注文日', '部署名', '氏名', 'カテゴリ', '商品内容', '合計金額(税込)']);

    foreach ($orders as $order) {
        $items = [];
        $cats  = [];
        foreach ($order->get_items() as $item) {
            $items[] = $item->get_name() . ' x' . $item->get_quantity();
            $terms = get_the_terms($item->get_product_id(), 'product_cat');
            if ($terms) {
                foreach ($terms as $term) {
                    $cats[] = $term->name;
                }
            }
        }
        fputcsv($output, [
            $order->get_order_number(),
            $order->get_date_created()->date('Y/m/d'),
            $order->get_billing_company(),
            $order->get_billing_last_name() . ' ' . $order->get_billing_first_name(),
            implode(' | ', array_unique($cats)),
            implode(' | ', $items),
            $order->get_total(),
        ]);
    }
    fclose($output);
    exit;
});
