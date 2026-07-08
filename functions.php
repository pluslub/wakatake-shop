<?php

defined( 'GTMT_VERSION' ) || define( 'GTMT_VERSION', wp_get_theme()->get( 'Version' ) );

include get_theme_file_path( '/inc/utils.php' );
include get_theme_file_path( '/inc/extension.php' );
include get_theme_file_path( '/inc/gutenmate-block-style.php' );
include get_theme_file_path( '/inc/html-processor.php' );
include get_theme_file_path( '/inc/health-check.php' );
include get_theme_file_path( '/inc/query-mod-use-taxonomy.php' );
include get_theme_file_path( '/inc/svg.php' );
include get_theme_file_path( '/inc/woocommerce.php' );
include get_theme_file_path( '/inc/compatibility.php' );
include get_theme_file_path( '/inc/theme-options/theme-options.php' );
include get_theme_file_path( '/inc/taxonomy-image-manager.php' );
include get_theme_file_path( '/plugin-activation/plugin-activation.php' );
include get_theme_file_path( '/demo/demo.php' );
include get_theme_file_path( '/build/blocks/gtmt--tab-group/index.php' );
include get_theme_file_path( '/build/blocks/gtmt--tab-panel/index.php' );
include get_theme_file_path( '/build/blocks/gtmt--product-image-slider/index.php' );
include get_theme_file_path( '/extension/block--aria.php' );
include get_theme_file_path( '/extension/block--binding.php' );
include get_theme_file_path( '/extension/block--cover.php' );
include get_theme_file_path( '/extension/block--dimensions.php' );
include get_theme_file_path( '/extension/block--dynamic-link.php' );
include get_theme_file_path( '/extension/block--hover-effect.php' );
include get_theme_file_path( '/extension/block--image-dimensions.php' );
include get_theme_file_path( '/extension/block--inline-css.php' );
include get_theme_file_path( '/extension/block--popup.php' );
include get_theme_file_path( '/extension/block--slider.php' );
include get_theme_file_path( '/extension/block--typography.php' );
include get_theme_file_path( '/extension/block--visibility.php' );
include get_theme_file_path( '/extension/core--button.php' );
include get_theme_file_path( '/extension/core--image.php' );
include get_theme_file_path( '/extension/core--navigation.php' );
include get_theme_file_path( '/extension/core--post-excerpt.php' );
include get_theme_file_path( '/extension/core--post-featured-image.php' );
include get_theme_file_path( '/extension/core--search.php' );
include get_theme_file_path( '/extension/core--social-link.php' );
include get_theme_file_path( '/extension/woocommerce--customer-account.php' );
include get_theme_file_path( '/extension/woocommerce--product-image.php' );

add_action( 'init', 'gtmt_unsupport_core_patterns' );
function gtmt_unsupport_core_patterns() {
	remove_theme_support( 'core-block-patterns' );

	$core_block_patterns = array(
		'core/query-standard-posts',
		'core/query-medium-posts',
		'core/query-small-posts',
		'core/query-grid-posts',
		'core/query-large-title-posts',
		'core/query-offset-posts',
	);

	foreach ( $core_block_patterns as $core_block_pattern ) {
		unregister_block_pattern( $core_block_pattern );
	}
}

add_action( 'init', 'gtmt_setup_assets' );
function gtmt_setup_assets() {
	wp_register_style( 'gutenmate-theme', get_template_directory_uri() . '/style.css', array(), GTMT_VERSION );
	wp_style_add_data( 'gutenmate-theme', 'rtl', 'replace' );

	if ( is_child_theme() ) {
		wp_register_style( 'gutenmate-child-theme', get_stylesheet_directory_uri() . '/style.css', array( 'gutenmate-theme' ), GTMT_VERSION );
	}
}

add_action( 'after_setup_theme', 'gtmt_setup_theme' );
function gtmt_setup_theme() {
	/* Load string translation */
	// Load from child theme
	load_theme_textdomain( wp_get_theme()->get( 'TextDomain' ), get_stylesheet_directory() . '/languages' );
	// Load from parent theme
	load_theme_textdomain( wp_get_theme()->get( 'TextDomain' ), get_template_directory() . '/languages' );

	/* Support wordpress features */
	add_theme_support( 'post-thumbnails' );

	// Avoid double title tag when using yoastseo
	if ( ! function_exists( 'YoastSEO' ) ) {
		add_theme_support( 'title-tag' );
	}

	add_theme_support( 'automatic-feed-links' );
	add_theme_support( 'post-formats', array( 'video', 'audio', 'gallery' ) );
	add_theme_support( 'responsive-embeds' );

	/* Support gutenberg features */
	add_theme_support( 'align-wide' );
	add_theme_support( 'custom-spacing' );
	add_theme_support( 'custom-units' );
	add_theme_support( 'appearance-tools' );
	add_editor_style();
	add_editor_style( 'style.css' );

	/* Support woocommerce features */
	add_theme_support( 'woocommerce', array(
		'thumbnail_image_width'         => 550,
		'gallery_thumbnail_image_width' => 150,
		'single_image_width'            => 768,
	) );

	/* Register block pattern categories */
	register_block_pattern_category( 'gtmt-migrations', array(
		'label' => esc_html__( 'Migrations', 'cusio' ),
	) );

	/* Register block styles */
	gtmt_register_block_style( 'core/code',
		array(
			array(
				'name'  => 'gtmt-dark',
				'label' => esc_html__( 'Dark', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/details',
		array(
			array(
				'name'  => 'gtmt-pill',
				'label' => esc_html__( 'Pill', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-plus',
				'label' => esc_html__( 'Plus', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/group',
		array(
			array(
				'name'  => 'gtmt-widget-area',
				'label' => esc_html__( 'Widget area', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-section-alt',
				'label' => esc_html__( 'Alt. section', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( array( 'core/list', 'core/categories' ),
		array(
			array(
				'name'  => 'gtmt-bullet',
				'label' => esc_html__( 'Bullet', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-check',
				'label' => esc_html__( 'Check', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-menu-tiny',
				'label' => esc_html__( 'Tiny menu', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-plus',
				'label' => esc_html__( 'Plus', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-minus',
				'label' => esc_html__( 'Minus', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/button',
		array(
			array(
				'name'  => 'gtmt-small',
				'label' => esc_html__( 'Small fill', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-arrow',
				'label' => esc_html__( 'Arrow', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-arrow-only',
				'label' => esc_html__( 'Arrow only', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-more',
				'label' => esc_html__( 'More', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-text-arrow',
				'label' => esc_html__( 'Text with arrow', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-text-back-arrow',
				'label' => esc_html__( 'Text with back arrow', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/columns',
		array(
			array(
				'name'  => 'gtmt-content-right-sidebar',
				'label' => esc_html__( 'Content with right sidebar', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-content-left-sidebar',
				'label' => esc_html__( 'Content with left sidebar', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/heading',
		array(
			array(
				'name'  => 'gtmt-number-badge',
				'label' => esc_html__( 'Number badge', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-shadow',
				'label' => esc_html__( 'Shadow', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/gallery',
		array(
			array(
				'name'  => 'gtmt-hero',
				'label' => esc_html__( 'Hero', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-vertical-masonry',
				'label' => esc_html__( 'Vertical masonry', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/image',
		array(
			array(
				'name'  => 'gtmt-featured-icon',
				'label' => esc_html__( 'Featured icon', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/loginout',
		array(
			array(
				'name'  => 'gtmt-icon-only',
				'label' => esc_html__( 'Icon only', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/media-text',
		array(
			array(
				'name'  => 'gtmt-hero',
				'label' => esc_html__( 'Hero', 'cusio' ),
			),

			array(
				'name'  => 'gtmt-dark-vertical',
				'label' => esc_html__( 'Dark vertical', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/navigation',
		array(
			array(
				'name'  => 'gtmt-primary',
				'label' => esc_html__( 'Primary', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-pill',
				'label' => esc_html__( 'Pill', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-text',
				'label' => esc_html__( 'Text', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/paragraph',
		array(
			array(
				'name'  => 'gtmt-heading',
				'label' => esc_html__( 'Heading', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-badge',
				'label' => esc_html__( 'Badge', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/post-author',
		array(
			array(
				'name'  => 'gtmt-vertical',
				'label' => esc_html__( 'Vertical', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/post-terms',
		array(
			array(
				'name'  => 'gtmt-badge',
				'label' => esc_html__( 'Badge', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-small',
				'label' => esc_html__( 'Small', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-large',
				'label' => esc_html__( 'Large', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/read-more',
		array(
			array(
				'name'  => 'gtmt-button',
				'label' => esc_html__( 'Button', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/separator',
		array(
			array(
				'name'  => 'gtmt-thin',
				'label' => esc_html__( 'Thin', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/search',
		array(
			array(
				'name'  => 'gtmt-outline',
				'label' => esc_html__( 'Outline', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/site-title',
		array(
			array(
				'name'  => 'gtmt-white',
				'label' => esc_html__( 'White', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'core/table',
		array(
			array(
				'name'  => 'gtmt-specs',
				'label' => esc_html__( 'Specs', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-image',
		array(
			array(
				'name'  => 'gtmt-faded',
				'label' => esc_html__( 'Faded white', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/catalog-sorting',
		array(
			array(
				'name'  => 'gtmt-icon-only',
				'label' => esc_html__( 'Icon only', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-image-gallery',
		array(
			array(
				'name'  => 'gtmt-side-thumbnails',
				'label' => esc_html__( 'Side thumbnails', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-faded-side-thumbnails',
				'label' => esc_html__( 'Faded side thumbnails', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-categories',
		array(
			array(
				'name'  => 'gtmt-dark',
				'label' => esc_html__( 'Dark', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-light',
				'label' => esc_html__( 'Light', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-price',
		array(
			array(
				'name'  => 'gtmt-badge',
				'label' => esc_html__( 'Badge', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-rating',
		array(
			array(
				'name'  => 'gtmt-small',
				'label' => esc_html__( 'Small', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/product-button',
		array(
			array(
				'name'  => 'gtmt-icon',
				'label' => esc_html__( 'Icon', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-secondary',
				'label' => esc_html__( 'Secondary', 'cusio' ),
			),
		)
	);

	gtmt_register_block_style( 'woocommerce/customer-account',
		array(
			array(
				'name'  => 'gtmt-themed',
				'label' => esc_html__( 'Themed', 'cusio' ),
			),
			array(
				'name'  => 'gtmt-button',
				'label' => esc_html__( 'Button', 'cusio' ),
			),
		)
	);
}

// Enqueue theme style before the WP global inline style
add_action( 'wp_enqueue_scripts', 'gtmt_enqueue_assets', 9 );
function gtmt_enqueue_assets() {
	wp_enqueue_style( 'gutenmate-theme' );

	if ( is_child_theme() ) {
		wp_enqueue_style( 'gutenmate-child-theme' );
	}
}

// Replace photoswipe skin of WooCommerce
add_action( 'wp_enqueue_scripts', 'gtmt_enqueue_photoswipe_skin', 11 );
function gtmt_enqueue_photoswipe_skin() {
	wp_deregister_style( 'photoswipe-default-skin' );
	wp_register_style( 'photoswipe-default-skin', get_template_directory_uri() . '/assets/css/photoswipe/default-skin/default-skin.min.css', array( 'photoswipe' ), GTMT_VERSION );
}

/**
 * Swiper default
 */
add_filter( 'gtmt_swiper_default_options', 'gtmt_swiper_default_options', 10, 2 );
function gtmt_swiper_default_options( $options, $block ) {
	$options['breakpoints']['0']['spaceBetween']    = '16';
	$options['breakpoints']['1120']['spaceBetween'] = '34';
	$options['keyboard']                            = array( 'enabled' => true );

	return $options;
}

/**
 * Add a custom template part area
 */
add_filter( 'default_wp_template_part_areas', 'gtmt_template_part_areas', 9 );
function gtmt_template_part_areas( array $areas ) {
	$areas[] = array(
		'area'        => 'sidebar',
		'area_tag'    => 'aside',
		'label'       => esc_html__( 'Sidebar', 'cusio' ),
		'description' => esc_html__( 'The Sidebar template defines a page area that typically contains widgets, social links, or any other combination of blocks.', 'cusio' ),
		'icon'        => 'sidebar',
	);

	return $areas;
}

/**
 * Display category image on category archive
 */
add_action( 'woocommerce_archive_description', 'gtmt_woocommerce_category_image', 2 );
function gtmt_woocommerce_category_image() {
	if ( is_product_category() ) {
		global $wp_query;
		$cat          = $wp_query->get_queried_object();
		$thumbnail_id = get_term_meta( $cat->term_id, 'thumbnail_id', true );
		$image        = wp_get_attachment_url( $thumbnail_id );
		if ( $image ) {
			echo '<img src="' . esc_url( $image ) . '" alt="' . esc_attr( $cat->name ) . '" />';
		}
	}
}
add_action( 'wp_footer', 'gtmt_billing_password_fix_script' );
function gtmt_billing_password_fix_script() {
    if ( function_exists('is_checkout') && is_checkout() ) {
        ?>
        <script type="text/javascript">
            jQuery(function($){
                function injectBillingPasswordDetails() {
                    // テーマ独自拡張のID「#billing_password」を取得
                    var $passwordInput = $('#billing_password');
                    
                    if ($passwordInput.length) {
                        // 親のラッパー要素（通常はpタグやdivタグ）を取得
                        var $fieldWrapper = $passwordInput.closest('.form-row') || $passwordInput.parent();
                        
                        // 1. ラベルが消えている、または空の場合に強制挿入
                        var $label = $fieldWrapper.find('label');
                        if ($label.length === 0 || $label.text().trim() === '') {
                            if($label.length) $label.remove();
                            $fieldWrapper.prepend('<label for="billing_password" style="display:block !important; font-weight:bold; margin-bottom:5px;">アカウント用パスワード <span class="required" style="color:red;">*</span></label>');
                        }

                        // 2. 入力欄の中のプレースホルダー（薄い文字）を設定
                        if (!$passwordInput.attr('placeholder') || $passwordInput.attr('placeholder') === '') {
                            $passwordInput.attr('placeholder', 'パスワードを入力してください');
                        }
                        
                        // 3. フィールド全体の直前に説明文を追加（重複防止付き）
                        if (!$('.custom-billing-pw-desc').length) {
                            $fieldWrapper.before('<p class="custom-billing-pw-desc" style="margin: 15px 0 5px; color: #333; font-weight: bold; font-size: 14px; width: 100%; display:block;">【会員登録パスワード】<br><span style="font-size: 12px; font-weight: normal; color: #666;">次回以降、マイアカウントにログインするためのパスワードを設定してください。</span></p>');
                        }
                    }
                }

                // 動的なReact書き換えにも対抗できるように監視（MutationObserver）
                var observer = new MutationObserver(function(mutations) {
                    injectBillingPasswordDetails();
                });

                var targetContainer = document.querySelector('.woocommerce-checkout') || document.body;
                observer.observe(targetContainer, {
                    childList: true,
                    subtree: true
                });

                // 初回実行
                injectBillingPasswordDetails();
            });
        </script>
        <?php
    }
}
/**
 * WooCommerce マイアカウント：文言の変更 ＆ サイドメニューの項目削除
 */
add_filter( 'gettext', 'custom_woocommerce_dashboard_text', 20, 3 );
function custom_woocommerce_dashboard_text( $translated_text, $text, $domain ) {
    if ( 'woocommerce' === $domain ) {
        
        // 1. 挨拶文の変更（※ %1$s や %2$s などの記号は消さずに残してください）
        if ( strpos( $text, 'Hello %1$s' ) !== false ) {
            $translated_text = '%1$s 様、こんにちは (アカウントが異なる場合は <a href="%2$s">ログアウト</a>)';
        }
        
        // 2. 説明文の変更
        if ( strpos( $text, 'From your account dashboard you can view' ) !== false ) {
            $translated_text = 'マイページへようこそ。こちらでは最近の注文履歴の確認、パスワードやアカウント詳細の編集が行えます。';
        }
        
    }
    return $translated_text;
}

add_filter( 'woocommerce_account_menu_items', 'custom_woocommerce_myaccount_menu_items', 99 );
function custom_woocommerce_myaccount_menu_items( $items ) {
    
    // 3. サイドバーから不要な項目を削除
    unset( $items['downloads'] ); // 「ダウンロード」を削除
    unset( $items['edit-address'] ); // 「住所」を削除
    
    return $items;
}
// 1. ダッシュボードに専用のウィジェット（枠）を登録する
add_action( 'wp_dashboard_setup', 'add_unordered_departments_dashboard_widget' );

function add_unordered_departments_dashboard_widget() {
    wp_add_dashboard_widget(
        'wc_unordered_departments_widget',
        '🛒 未注文部署（グループ）一覧',
        'display_unordered_departments_dashboard_widget'
    );
}

// 2. ウィジェットの中身を表示する関数
function display_unordered_departments_dashboard_widget() {
    // --- 【設定】HTMLから抽出した正確な全55部署マスターリスト ---
    $master_departments = array(
        "製造第1課 橋本 竜一",
        "製造第1課 勝見 久美",
        "製造第1課 木戸 靖典",
        "製造第1課 左山 裕太",
        "製造第1課 居垣 勝彦",
        "製造第1課 斉藤 秋弥",
        "製造第2課 Ｆ１ライン",
        "製造第2課 Ｆ２ライン",
        "製造第2課 Ｆ３ライン",
        "２SVライン",
        "製造第2課 樹脂組立",
        "製造第2課 ２ＡＸライン",
        "製造第2課 T熱交",
        "製造第2課 Ｓ２ライン",
        "製造第2課 ＦＣ・Ｆ８ライン",
        "製造第2課 1物流",
        "製造第2課 2物流",
        "製造第2課 3物流",
        "製造第3課 井川クロス・ＯＤＭ",
        "製造第3課 井川樹脂加工",
        "製造第3課 井川Ｋ2・ストリーマライン",
        "製造第3課 信藤Ｆ６ライン加工 １Ｆ",
        "製造第3課 信藤Ｆ６ライン組立 3F",
        "製造第3課 信藤３ＣＸライン",
        "製造第3課 信藤ＲＡ・ＰＡサービス",
        "製造第3課 亀尾ＦＴライン",
        "製造第3課 亀尾ＦＢライン",
        "製造第3課 亀尾ＦＷライン",
        "製造第3課 亀尾Ｆ７ライン",
        "滋賀製造部 武田 泰樹",
        "滋賀製造部 市村 裕章",
        "滋賀製造部 木内 慎弥",
        "品質管理部 德原 紗英 宇野 智久",
        "物流本部 三木 勇太",
        "サービス本部 高崎 竜河",
        "滋賀製作所 澁谷 美沙子",
        "ダイキン福祉サービス 梅景 佳則",
        "生産技術部 松下 典正",
        "生産技術部 村上 亮太",
        "生産技術部 蛭子 奉紀",
        "生産技術部 山本 大貴",
        "グローバル事業推進部 生技センター 孔 碧エイ",
        "SCM部 深畑 聡子",
        "企画部 平田 将太郎",
        "グローバル調達本部 幣 雄一朗",
        "DKR 杉山 智美",
        "DKP 寺尾 正人",
        "モータグループ 兒玉 光平",
        "小型RA商品グループ 山寺 哲史 神棒 圭太",
        "住宅設備商品グループ 上別府 和熙 及川 学",
        "住宅用空気商品グループ 米女 一",
        "開発企画グループ 北川 和也",
        "先行要素・基盤技術グループ 吾郷 祥太",
        "デバイス技術グループ 岸田 匠人 恒岡 知生",
        "圧縮機グループ 砂原 裕也",
        "開発信頼性グループ 青森壮 佐野 久瑠美",
        "組合専従 組合"
    );

    // 過去すべての注文データを取得（一度でも購入したことがある部署を判定）
    $orders = wc_get_orders( array(
        'limit'  => -1,
        'status' => array( 'wc-processing', 'wc-completed', 'wc-on-hold' ),
    ) );

    // 実際に注文があった部署名（_billing_company）を抽出
    $ordered_departments = array();
    foreach ( $orders as $order ) {
        $company_or_dept = $order->get_billing_company(); 
        if ( ! empty( $company_or_dept ) ) {
            $ordered_departments[] = trim( $company_or_dept );
        }
    }
    $ordered_departments = array_unique( $ordered_departments );

    // 全部署から注文済みを差し引いて「未注文」を割り出す
    $unordered_departments = array_diff( $master_departments, $ordered_departments );

    // --- 画面への出力 HTML ---
    echo '<div class="wc-unordered-deps-wrapper" style="max-height: 400px; overflow-y: auto;">';
    
    if ( ! empty( $unordered_departments ) ) {
        echo '<p style="color: #d63638; font-weight: bold; margin-bottom: 10px;">⚠️ まだ注文がない部署・担当（' . count( $unordered_departments ) . '）</p>';
        echo '<ul style="list-style-type: disc; padding-left: 20px; margin-bottom: 20px;">';
        foreach ( $unordered_departments as $dept ) {
            echo '<li style="color: #d63638; font-size: 13px; margin-bottom: 4px;">' . esc_html( $dept ) . '</li>';
        }
        echo '</ul>';
    } else {
        echo '<p style="color: #00a32a; font-weight: bold; font-size: 14px; background: #ecf7ed; padding: 10px; border-left: 4px solid #00a32a;">🎉 すべての部署が今月注文済みです！</p>';
    }

    // （参考）注文が完了している部署を表示
    echo '<details style="margin-top: 15px; border-top: 1px solid #eee; padding-top: 10px;">';
    echo '<summary style="cursor: pointer; color: #646970; font-size: 12px;">注文済みの部署を表示する (' . count( $ordered_departments ) . ')</summary>';
    if ( ! empty( $ordered_departments ) ) {
        echo '<ul style="list-style-type: circle; padding-left: 20px; margin-top: 5px; color: #646970; font-size: 12px;">';
        foreach ( $ordered_departments as $dept ) {
            echo '<li>' . esc_html( $dept ) . '</li>';
        }
        echo '</ul>';
    } else {
        echo '<p style="font-size: 12px; color: #646970; margin-top: 5px;">まだどの部署も注文していません。</p>';
    }
    echo '</details>';
    
    echo '</div>';
}