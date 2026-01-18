<?php
/**
 * Class to handle WooCommerce to Google Sheets integration
 */
namespace UGSIW;
use UGSIW\UGSIW_Script_Generator;

if (!defined('ABSPATH')) {
    exit;
}

class UGSIW_To_Google_Sheets {
    
    private $available_fields = array();
    private $is_pro_active = false;
    
    public function __construct() {
        // Check if Pro version is active
        $this->is_pro_active = apply_filters('is_active_ultimate_ugsiw_pro_feature', false);
        
        // Define available fields
        $this->define_available_fields();
        
        // Check if WooCommerce is active
        add_action('admin_init', array($this, 'wpmethods_check_woocommerce'));

        // Initialize script generator
        $this->script_generator = new UGSIW_Script_Generator($this->available_fields);
        
        // Hook into order status changes
        add_action('woocommerce_order_status_changed', array($this, 'wpmethods_send_order_to_sheets'), 10, 4);
        
        // Add settings page
        add_action('admin_menu', array($this, 'wpmethods_add_admin_menu'));
        add_action('admin_init', array($this, 'wpmethods_settings_init'));
        
        // Add admin scripts
        add_action('admin_enqueue_scripts', array($this, 'wpmethods_admin_scripts'));
        
        // AJAX handler for generating Google Apps Script
        add_action('wp_ajax_wpmethods_generate_google_script', array($this, 'wpmethods_generate_google_script_ajax'));
        
        // Register activation hook to set default values
        register_activation_hook(__FILE__, array($this, 'wpmethods_activate_plugin'));
    }
    
    /**
     * Define available fields
     */
    private function define_available_fields() {
        // Simplified fields - only include working ones
        $this->available_fields = array(
            'order_id' => array(
                'label' => 'Order ID',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-cart'
            ),
            'order_amount_with_currency' => array(
                'label' => 'Order Amount',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-money'
            ),
            'order_currency' => array(
                'label' => 'Order Currency',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-tag'
            ),
            'billing_name' => array(
                'label' => 'Billing Name',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-admin-users'
            ),
            'billing_email' => array(
                'label' => 'Email Address',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-email'
            ),
            'billing_phone' => array(
                'label' => 'Phone',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-phone'
            ),
            'billing_address' => array(
                'label' => 'Billing Address',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-location'
            ),
            'product_name' => array(
                'label' => 'Product Name',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-products'
            ),
            
            'payment_method_title' => array(
                'label' => 'Payment Method',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-bank'
            ),
            'order_status' => array(
                'label' => 'Order Status',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-yes'
            ),
            'order_date' => array(
                'label' => 'Order Date',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-calendar'
            ),
            'product_categories' => array(
                'label' => 'Product Categories',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-category'
            ),
            'customer_note' => array(
                'label' => 'Customer Note',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-edit'
            ),
            'shipping_method' => array(
                'label' => 'Shipping Method',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-car'
            ),
            'product_type' => array(
                'label' => 'Product Type',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-tag'
            )
        );
        
        // Add pro fields (always include so non-active users can see PRO badges)
        $this->available_fields = array_merge($this->available_fields, array(
            'customer_id' => array(
                'label' => 'Customer ID',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-id',
                'pro' => true
            ),
            'coupon_used' => array(
                'label' => 'Coupon Used',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-tag',
                'pro' => true
            ),
            'product_sku' => array(
                'label' => 'Product SKU',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-welcome-widgets-menus',
                'pro' => true
            ),
            'product_quantity' => array(
                'label' => 'Product Quantity',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-editor-ol',
                'pro' => true
            ),
            'product_price' => array(
                'label' => 'Product Price',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-money-alt',
                'pro' => true
            ),
            'product_total' => array(
                'label' => 'Product Total',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-calculator',
                'pro' => true
            )
        ));
    }
    
    /**
     * Get selected fields for Google Sheets
     */
    private function wpmethods_get_selected_fields() {
        $selected_fields = get_option('ugsiw_gs_selected_fields', array());
        
        // Ensure it's always an array
        if (!is_array($selected_fields)) {
            if (is_string($selected_fields) && !empty($selected_fields)) {
                $unserialized = maybe_unserialize($selected_fields);
                if (is_array($unserialized)) {
                    $selected_fields = $unserialized;
                } else {
                    $selected_fields = array_map('trim', explode(',', $selected_fields));
                }
            } else {
                $selected_fields = array();
            }
        }
        
        // Always include required fields
        foreach ($this->available_fields as $field_key => $field_info) {
            if (isset($field_info['always_include']) && $field_info['always_include']) {
                if (!in_array($field_key, $selected_fields)) {
                    $selected_fields[] = $field_key;
                }
            }
        }
        
        return array_unique($selected_fields);
    }
    
    /**
     * Check if WooCommerce is active
     */
    public function wpmethods_check_woocommerce() {
        if (!class_exists('WooCommerce')) {
            add_action('admin_notices', array($this, 'wpmethods_woocommerce_missing_notice'));
        }
    }
    
    /**
     * WooCommerce missing notice
     */
    public function wpmethods_woocommerce_missing_notice() {
        ?>
        <div class="notice notice-error">
            <p><?php esc_html_e('WP Methods WooCommerce to Google Sheets requires WooCommerce to be installed and activated.', 'ultimate-google-sheets-integration-for-woo'); ?></p>
        </div>
        <?php
    }
    
    /**
     * Plugin activation - set default values
     */
    public function wpmethods_activate_plugin() {
        if (!get_option('ugsiw_gs_order_statuses')) {
            update_option('ugsiw_gs_order_statuses', array('completed', 'processing'));
        }
        
        if (!get_option('ugsiw_gs_script_url')) {
            update_option('ugsiw_gs_script_url', '');
        }
        
        if (!get_option('ugsiw_gs_product_categories')) {
            update_option('ugsiw_gs_product_categories', array());
        }
        
        // Set default selected fields (all fields)
        if (!get_option('ugsiw_gs_selected_fields')) {
            $default_fields = array_keys($this->available_fields);
            update_option('ugsiw_gs_selected_fields', $default_fields);
        }
    }
    
    /**
     * Get all WooCommerce order statuses
     */
    private function wpmethods_get_wc_order_statuses() {
        $statuses = wc_get_order_statuses();
        $clean_statuses = array();
        
        foreach ($statuses as $key => $label) {
            $clean_key = str_replace('wc-', '', $key);
            $clean_statuses[$clean_key] = $label;
        }
        
        return $clean_statuses;
    }
    
    /**
     * Get selected categories as array
     */
    private function wpmethods_get_selected_categories() {
        $selected_categories = get_option('ugsiw_gs_product_categories', array());
        
        if (!is_array($selected_categories)) {
            if (is_string($selected_categories) && !empty($selected_categories)) {
                $unserialized = maybe_unserialize($selected_categories);
                if (is_array($unserialized)) {
                    $selected_categories = $unserialized;
                } else {
                    $selected_categories = array_map('trim', explode(',', $selected_categories));
                }
            } else {
                $selected_categories = array();
            }
        }
        
        $selected_categories = array_map('intval', $selected_categories);
        
        return $selected_categories;
    }

    /**
     * Get product types from order
     */
    private function get_order_product_types($order) {
        $product_types = array();
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                $product_type = $product->get_type();
                
                // Map WooCommerce product types to readable names
                $type_mapping = array(
                    'simple' => 'Simple',
                    'variable' => 'Variable',
                    'variation' => 'Variation',
                    'grouped' => 'Grouped',
                    'external' => 'External/Affiliate',
                    'subscription' => 'Subscription',
                    'variable-subscription' => 'Variable Subscription',
                    'booking' => 'Booking'
                );
                
                $readable_type = isset($type_mapping[$product_type]) 
                    ? $type_mapping[$product_type] 
                    : ucfirst($product_type);
                    
                $product_types[] = $readable_type;
            } else {
                // Product might be deleted
                $product_types[] = 'Unknown/Deleted';
            }
        }
        
        $product_types = array_unique($product_types);
        
        return !empty($product_types) ? implode(', ', $product_types) : 'Simple';
    }
    
    /**
     * Check if order contains products from selected categories
     */
    private function wpmethods_order_has_selected_categories($order, $selected_categories) {
        if (empty($selected_categories)) {
            return true;
        }
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                // Determine correct product ID (handle variations)
                $product_id = $product->get_id();
                if (method_exists($product, 'is_type') && $product->is_type('variation')) {
                    // Prefer parent ID for variations
                    if (method_exists($product, 'get_parent_id')) {
                        $parent_id = $product->get_parent_id();
                        if ($parent_id) {
                            $product_id = $parent_id;
                        }
                    } elseif (method_exists($item, 'get_variation_id')) {
                        $variation_parent = $item->get_variation_id();
                        if ($variation_parent) {
                            $product_id = $variation_parent;
                        }
                    }
                }

                $product_categories = wp_get_post_terms($product_id, 'product_cat', array('fields' => 'ids'));

                if (!empty(array_intersect($product_categories, $selected_categories))) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * Send order data to Google Sheets with retry mechanism
     */
    public function wpmethods_send_order_to_sheets($order_id, $old_status, $new_status, $order) {
        
        $selected_statuses = get_option('ugsiw_gs_order_statuses', array('completed', 'processing'));
        
        if (!is_array($selected_statuses)) {
            $selected_statuses = maybe_unserialize($selected_statuses);
            if (!is_array($selected_statuses)) {
                $selected_statuses = array('completed', 'processing');
            }
        }
        
        if (!in_array($new_status, $selected_statuses)) {
            return;
        }
        
        $selected_categories = $this->wpmethods_get_selected_categories();
        
        if (!$this->wpmethods_order_has_selected_categories($order, $selected_categories)) {
            return;
        }
        
        $order_data = $this->wpmethods_prepare_order_data($order);

        // If Pro forwarding is enabled, attempt to POST to the webhook even if Apps Script failed
        if ($this->is_pro_active && get_option('ugsiw_gs_forward_webhook', '0') === '1') {
            $webhook = get_option('ugsiw_gs_webhook_url', '');
            if (!empty($webhook)) {
                $hook_response = wp_remote_post($webhook, array(
                    'method' => 'POST',
                    'timeout' => 10,
                    'redirection' => 2,
                    'httpversion' => '1.1',
                    'blocking' => true,
                    'headers' => array('Content-Type' => 'application/json'),
                    'body' => wp_json_encode($order_data),
                    'data_format' => 'body'
                ));

                if (is_wp_error($hook_response)) {
                    error_log('UGSIW Webhook Forward Error (fallback): ' . $hook_response->get_error_message());
                } else {
                    $hook_code = wp_remote_retrieve_response_code($hook_response);
                    if ($hook_code !== 200 && $hook_code !== 201) {
                        $hook_body = wp_remote_retrieve_body($hook_response);
                        $hook_clean = wp_strip_all_tags($hook_body);
                        if (is_string($hook_clean) && strlen($hook_clean) > 200) {
                            $hook_clean = substr($hook_clean, 0, 200) . '...';
                        }
                    }
                }
            }
        }

        // Get Apps Script URL
        $script_url = get_option('ugsiw_gs_script_url', '');

        if (empty($script_url)) {
            return;
        }
        
        // Add retry mechanism for failed requests
        $max_retries = 2;
        $retry_delay = 1; // seconds
        
        for ($attempt = 1; $attempt <= $max_retries; $attempt++) {
            $response = wp_remote_post($script_url, array(
                'method' => 'POST',
                'timeout' => 10,
                'redirection' => 2,
                'httpversion' => '1.1',
                'blocking' => true,
                'headers' => array(
                    'Content-Type' => 'application/json',
                ),
                'body' => wp_json_encode($order_data),
                'data_format' => 'body'
            ));
            
            // Check if request was successful
            if (!is_wp_error($response)) {
                $response_code = wp_remote_retrieve_response_code($response);
                if ($response_code === 200) {
                    // Success
                    $body = wp_remote_retrieve_body($response);
                    $data = json_decode($body, true);
                    
                    if (isset($data['status']) && $data['status'] === 'success') {
                        return; // Exit on success
                    }
                } else {
                    $body = wp_remote_retrieve_body($response);
                    // Keep logs concise: strip HTML and truncate long responses
                    $clean_body = wp_strip_all_tags($body);
                    if (is_string($clean_body) && strlen($clean_body) > 200) {
                        $clean_body = substr($clean_body, 0, 200) . '...';
                    }
                }
            } else {
                // Log WP error
                error_log('UGSIW Google Sheets Error: ' . $response->get_error_message());
            }
            
            // If not last attempt, wait and retry
            if ($attempt < $max_retries) {
                sleep($retry_delay * $attempt); // Exponential backoff
            }
        }
        
        
    }
    
    /**
     * Prepare order data for Google Sheets based on selected fields
     */
    private function wpmethods_prepare_order_data($order) {
        $selected_fields = $this->wpmethods_get_selected_fields();
        $order_data = array();
        
        // If Pro is not active, ensure pro-only fields are not included in the payload
        if (!$this->is_pro_active) {
            foreach ($this->available_fields as $fkey => $finfo) {
                if (isset($finfo['pro']) && $finfo['pro']) {
                    if (($k = array_search($fkey, $selected_fields)) !== false) {
                        unset($selected_fields[$k]);
                    }
                }
            }
            $selected_fields = array_values($selected_fields);
        }
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $value = $this->get_field_value($field_key, $order);
                if ($value !== null) {
                    $order_data[$field_key] = $value;
                }
            }
        }
        
        return $order_data;
    }
    
    /**
     * Get value for a specific field
     */
    private function get_field_value($field_key, $order) {
        switch ($field_key) {
            case 'order_id':
                return $order->get_id();

            case 'order_amount_with_currency':
                $currency_symbol = $order->get_currency();
                $currency_symbol_formatted = get_woocommerce_currency_symbol($currency_symbol);
                // Fix: Don't use HTML entities, just send the raw symbol
                $symbol = $currency_symbol_formatted;
                // If it's an HTML entity, decode it
                if (strpos($symbol, '&#') !== false) {
                    $symbol = html_entity_decode($symbol, ENT_QUOTES, 'UTF-8');
                }
                return $symbol . $order->get_total();
                
            case 'order_currency':
                return $order->get_currency();
                
                
            case 'billing_name':
                return $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
                
            case 'billing_email':
                return $order->get_billing_email();
                
            case 'billing_phone':
                return $order->get_billing_phone();
                
            case 'billing_address':
                $address_parts = array();
                
                if ($address1 = $order->get_billing_address_1()) {
                    $address_parts[] = $address1;
                }
                if ($address2 = $order->get_billing_address_2()) {
                    $address_parts[] = $address2;
                }
                if ($city = $order->get_billing_city()) {
                    $address_parts[] = $city;
                }
                if ($state = $order->get_billing_state()) {
                    $address_parts[] = $state;
                }
                if ($postcode = $order->get_billing_postcode()) {
                    $address_parts[] = $postcode;
                }
                if ($country = $order->get_billing_country()) {
                    $address_parts[] = $country;
                }
                
                return implode(', ', $address_parts);
                
            case 'product_name':
                $product_names = array();
                foreach ($order->get_items() as $item) {
                    $product_names[] = $item->get_name();
                }
                return implode(', ', $product_names);
                
            case 'payment_method_title':
                $payment_title = $order->get_payment_method_title();
                if (empty($payment_title)) {
                    // Try to get payment gateway title
                    $payment_method = $order->get_payment_method();
                    if ($payment_method) {
                        $gateways = WC()->payment_gateways()->payment_gateways();
                        if (isset($gateways[$payment_method])) {
                            $payment_title = $gateways[$payment_method]->get_title();
                        }
                    }
                }
                return $payment_title ? $payment_title : '';
                
            case 'order_status':
                return $order->get_status();
                
            case 'order_date':
                $date_created = $order->get_date_created();
                return $date_created ? $date_created->format('Y-m-d H:i:s') : '';
                
            case 'product_categories':
                return $this->wpmethods_get_order_categories($order);

            case 'coupon_used':
                // Try WC_Order::get_used_coupons() first
                if (method_exists($order, 'get_used_coupons')) {
                    $coupons = $order->get_used_coupons();
                    if (is_array($coupons) && !empty($coupons)) {
                        return implode(', ', $coupons);
                    }
                }

                // Fallback: inspect coupon items
                $coupon_codes = array();
                foreach ($order->get_items('coupon') as $coupon_item) {
                    if (is_array($coupon_item) && isset($coupon_item['code'])) {
                        $coupon_codes[] = $coupon_item['code'];
                    } elseif (is_object($coupon_item) && method_exists($coupon_item, 'get_code')) {
                        $coupon_codes[] = $coupon_item->get_code();
                    }
                }

                $coupon_codes = array_filter($coupon_codes, function($c) { return trim($c) !== ''; });
                return !empty($coupon_codes) ? implode(', ', $coupon_codes) : '';

            case 'product_sku':
                $skus = array();
                foreach ($order->get_items() as $item) {
                    $product = $item->get_product();
                    if ($product && method_exists($product, 'get_sku')) {
                        $sku = $product->get_sku();
                        if (!empty($sku)) {
                            $skus[] = $sku;
                        }
                    } else {
                        $skus[] = '';
                    }
                }
                $skus = array_filter($skus, function($v) { return $v !== ''; });
                return !empty($skus) ? implode(', ', $skus) : '';

            case 'product_quantity':
                $quantities = array();
                foreach ($order->get_items() as $item) {
                    $qty = $item->get_quantity();
                    $quantities[] = $qty;
                }
                return !empty($quantities) ? implode(', ', $quantities) : '';

            case 'product_price':
                $prices = array();
                $currency = $order->get_currency();
                $currency_symbol = get_woocommerce_currency_symbol($currency);
                // Decode HTML entities like &#2547; and remove non-breaking spaces
                $currency_symbol = html_entity_decode($currency_symbol, ENT_QUOTES, 'UTF-8');
                $currency_symbol = str_replace("\xc2\xa0", ' ', $currency_symbol);
                $decimals = function_exists('wc_get_price_decimals') ? wc_get_price_decimals() : 2;
                foreach ($order->get_items() as $item) {
                    $qty = max(1, (int) $item->get_quantity());
                    $per_item_price = $qty ? ($item->get_total() / $qty) : $item->get_total();
                    $formatted = number_format((float) $per_item_price, $decimals, '.', '');
                    $prices[] = $currency_symbol . $formatted;
                }
                return !empty($prices) ? implode(', ', $prices) : '';

            case 'product_total':
                $totals = array();
                $currency = $order->get_currency();
                $currency_symbol = get_woocommerce_currency_symbol($currency);
                $currency_symbol = html_entity_decode($currency_symbol, ENT_QUOTES, 'UTF-8');
                $currency_symbol = str_replace("\xc2\xa0", ' ', $currency_symbol);
                $decimals = function_exists('wc_get_price_decimals') ? wc_get_price_decimals() : 2;
                foreach ($order->get_items() as $item) {
                    $formatted = number_format((float) $item->get_total(), $decimals, '.', '');
                    $totals[] = $currency_symbol . $formatted;
                }
                return !empty($totals) ? implode(', ', $totals) : '';

            case 'product_type':
                return $this->get_order_product_types($order);
            
            case 'customer_id':
                // Return the WP user ID for the customer if available, otherwise empty
                $cust_id = null;
                if (method_exists($order, 'get_customer_id')) {
                    $cust_id = $order->get_customer_id();
                } elseif (method_exists($order, 'get_user_id')) {
                    $cust_id = $order->get_user_id();
                }
                return $cust_id ? (string) $cust_id : '';

            case 'customer_note':
                return $order->get_customer_note();

            case 'shipping_method':
                $shipping_entries = array();
                $currency = $order->get_currency();
                $currency_symbol = get_woocommerce_currency_symbol($currency);
                $currency_symbol = html_entity_decode($currency_symbol, ENT_QUOTES, 'UTF-8');
                $currency_symbol = str_replace("\xc2\xa0", ' ', $currency_symbol);
                $decimals = function_exists('wc_get_price_decimals') ? wc_get_price_decimals() : 2;
                foreach ($order->get_shipping_methods() as $shipping_item) {
                    $name = '';
                    if (method_exists($shipping_item, 'get_method_title') && !empty($shipping_item->get_method_title())) {
                        $name = $shipping_item->get_method_title();
                    } elseif (method_exists($shipping_item, 'get_name') && !empty($shipping_item->get_name())) {
                        $name = $shipping_item->get_name();
                    } else {
                        $name = method_exists($order, 'get_shipping_method') ? $order->get_shipping_method() : '';
                    }
                    $total = 0;
                    if (method_exists($shipping_item, 'get_total')) {
                        $total = (float) $shipping_item->get_total();
                    }
                    if ($total !== 0) {
                        $formatted = number_format($total, $decimals, '.', '');
                        $shipping_entries[] = trim($name . ' - ' . $currency_symbol . $formatted);
                    } else {
                        $shipping_entries[] = trim($name);
                    }
                }
                return !empty($shipping_entries) ? implode(', ', $shipping_entries) : '';
                
            default:
                return null;
        }
    }
    
    /**
     * Get categories from order products
     */
    private function wpmethods_get_order_categories($order) {
        $all_categories = array();
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            if ($product) {
                // Resolve parent/product id for variations
                $product_id = $product->get_id();
                if (method_exists($product, 'is_type') && $product->is_type('variation')) {
                    if (method_exists($product, 'get_parent_id')) {
                        $parent_id = $product->get_parent_id();
                        if ($parent_id) {
                            $product_id = $parent_id;
                        }
                    } elseif (method_exists($item, 'get_variation_id')) {
                        $variation_parent = $item->get_variation_id();
                        if ($variation_parent) {
                            $product_id = $variation_parent;
                        }
                    }
                }

                $categories = wp_get_post_terms($product_id, 'product_cat', array('fields' => 'names'));
                if (is_array($categories) && !empty($categories)) {
                    $all_categories = array_merge($all_categories, $categories);
                }
            }
        }
        
        $all_categories = array_unique($all_categories);
        
        return implode(', ', $all_categories);
    }
    
    /**
     * Add admin menu
     */
    public function wpmethods_add_admin_menu() {
        add_menu_page(
            'WooCommerce to Google Sheets',
            'WC Orders to Google Sheets',
            'manage_options',
            'wpmethods-wc-to-google-sheets',
            array($this, 'wpmethods_settings_page'),
            'dashicons-google'
        );
    }
    
    /**
     * Admin scripts
     */
    public function wpmethods_admin_scripts($hook) {
        if ($hook != 'toplevel_page_wpmethods-wc-to-google-sheets') {
            return;
        }

        $min = defined('SCRIPT_DEBUG') && SCRIPT_DEBUG ? '' : '.min';
        
        // Enqueue WordPress core styles
        wp_enqueue_style('wp-color-picker');
        wp_enqueue_script('wp-color-picker');
        wp_enqueue_style('wpmethods-wc-gs-admin-style', plugin_dir_url(__FILE__) . '../assets/css/style' . $min . '.css', array(), UGSIW_VERSION);
        wp_enqueue_script('wpmethods-wc-gs-admin-script', plugin_dir_url(__FILE__) . '../assets/js/admin' . $min . '.js', array('jquery'), UGSIW_VERSION, true);
       
        // Localize script to pass PHP variables to JS
        wp_localize_script('wpmethods-wc-gs-admin-script', 'ugsiw_gs', array(
            'ajax_url' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('wpmethods_generate_script_nonce')
        ));
    }
    
    /**
     * Initialize settings
     */
    public function wpmethods_settings_init() {
        // Main settings
        register_setting('ugsiw_gs_settings', 'ugsiw_gs_order_statuses', array($this, 'wpmethods_sanitize_array'));
        register_setting('ugsiw_gs_settings', 'ugsiw_gs_script_url', 'esc_url_raw');
        register_setting('ugsiw_gs_settings', 'ugsiw_gs_product_categories', array($this, 'wpmethods_sanitize_array'));
        register_setting('ugsiw_gs_settings', 'ugsiw_gs_selected_fields', array($this, 'wpmethods_sanitize_array'));
        register_setting('ugsiw_gs_settings', 'ugsiw_gs_monthly_sheets', array($this, 'wpmethods_sanitize_checkbox'));
        
        // Pro settings (only if active)
        if ($this->is_pro_active) {
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_sheet_mode', array($this, 'wpmethods_sanitize_text'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_daily_weekly', array($this, 'wpmethods_sanitize_text'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_product_sheets', array($this, 'wpmethods_sanitize_checkbox'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_custom_sheet_name', array($this, 'wpmethods_sanitize_text'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_custom_name_template', array($this, 'wpmethods_sanitize_text'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_forward_webhook', array($this, 'wpmethods_sanitize_checkbox'));
            register_setting('ugsiw_gs_settings', 'ugsiw_gs_webhook_url', array($this, 'wpmethods_sanitize_text'));
        }
        
        // Main settings section
        add_settings_section(
            'ugsiw_gs_section',
            'Google Sheets Integration Settings',
            array($this, 'wpmethods_section_callback'),
            'ugsiw_gs_settings'
        );
        
        add_settings_field(
            'ugsiw_gs_order_statuses',
            'Trigger Order Statuses',
            array($this, 'wpmethods_order_statuses_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );
        
        // Replace separate monthly/daily/product/custom fields with a single Sheet Mode selector
        add_settings_field(
            'ugsiw_gs_sheet_mode',
            'Sheet Mode',
            array($this, 'wpmethods_sheet_mode_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );
        
        add_settings_field(
            'ugsiw_gs_product_categories',
            'Product Categories Filter',
            array($this, 'wpmethods_product_categories_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );

        add_settings_field(
            'ugsiw_gs_selected_fields',
            'Checkout Fields',
            array($this, 'wpmethods_selected_fields_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );
        
        add_settings_field(
            'ugsiw_gs_script_url',
            'Google Apps Script URL',
            array($this, 'wpmethods_script_url_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );

        // Pro webhook forwarding field
        add_settings_field(
            'ugsiw_gs_forward_webhook',
            'Forward Orders to Webhook (Pro)',
            array($this, 'wpmethods_forward_webhook_render'),
            'ugsiw_gs_settings',
            'ugsiw_gs_section'
        );
    }


    /**
     * Sanitize text input
     */
    public function wpmethods_sanitize_text($input) {
        return sanitize_text_field($input);
    }
    
    /**
     * Daily/Weekly Sheets field render - PRO FEATURE
     */
    public function wpmethods_daily_weekly_render() {
        if (!$this->is_pro_active) {
            echo '<div style="padding: 20px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">';
            echo '<p style="margin: 0; color: #856404; font-weight: 600;">';
            echo '<span class="dashicons dashicons-lock" style="color: #ffc107;"></span> ';
            echo 'This is a Pro feature. <a href="admin.php?page=ugsiw-license" style="color: #856404; text-decoration: underline;">Upgrade to Pro</a> to unlock Daily/Weekly Sheets.';
            echo '</p>';
            echo '</div>';
            return;
        }
        
        $value = get_option('ugsiw_gs_daily_weekly', 'none');
        ?>
        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
            <label style="display: flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_daily_weekly" value="none" <?php checked($value, 'none'); ?>>
                <span style="font-weight: 500;">
                    <span class="dashicons dashicons-no" style="color: #dc3545;"></span>
                    None
                </span>
            </label>
            
            <label style="display: flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_daily_weekly" value="daily" <?php checked($value, 'daily'); ?>>
                <span style="font-weight: 500;">
                    <span class="dashicons dashicons-calendar" style="color: #28a745;"></span>
                    Daily Sheets
                </span>
            </label>
            
            <label style="display: flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_daily_weekly" value="weekly" <?php checked($value, 'weekly'); ?>>
                <span style="font-weight: 500;">
                    <span class="dashicons dashicons-calendar-alt" style="color: #007bff;"></span>
                    Weekly Sheets
                </span>
            </label>
        </div>
        <p class="description" style="margin-top: 10px;">
            <span class="dashicons dashicons-star-filled" style="color: #ffc107;"></span>
            Pro Feature: Automatically create new sheets daily or weekly. Orders will be organized by date.
        </p>
        <?php
    }


    /**
     * Product-wise Sheets field render - PRO FEATURE
     */
    public function wpmethods_product_sheets_render() {
        if (!$this->is_pro_active) {
            echo '<div style="padding: 20px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">';
            echo '<p style="margin: 0; color: #856404; font-weight: 600;">';
            echo '<span class="dashicons dashicons-lock" style="color: #ffc107;"></span> ';
            echo 'This is a Pro feature. <a href="admin.php?page=ugsiw-license" style="color: #856404; text-decoration: underline;">Upgrade to Pro</a> to unlock Product-wise Sheets.';
            echo '</p>';
            echo '</div>';
            return;
        }
        
        $value = get_option('ugsiw_gs_product_sheets', '0');
        ?>
        <label style="display: inline-flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
            <input type="checkbox" name="ugsiw_gs_product_sheets" value="1" <?php checked($value, '1'); ?> style="width: 20px; height: 20px;">
            <span style="font-weight: 500; font-size: 14px;">
                <span class="dashicons dashicons-category" style="color: #667eea;"></span>
                Enable Product Category Sheets
            </span>
        </label>
        <p class="description" style="margin-top: 10px;">
            <span class="dashicons dashicons-star-filled" style="color: #ffc107;"></span>
            Pro Feature: Create separate sheets for each product category. Orders will be sorted into their respective category sheets.
        </p>
        <?php
    }
    
    /**
     * Custom Sheet Naming field render - PRO FEATURE
     */
    public function wpmethods_custom_sheet_name_render() {
        if (!$this->is_pro_active) {
            echo '<div style="padding: 20px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">';
            echo '<p style="margin: 0; color: #856404; font-weight: 600;">';
            echo '<span class="dashicons dashicons-lock" style="color: #ffc107;"></span> ';
            echo 'This is a Pro feature. <a href="admin.php?page=ugsiw-license" style="color: #856404; text-decoration: underline;">Upgrade to Pro</a> to unlock Custom Sheet Naming.';
            echo '</p>';
            echo '</div>';
            return;
        }
        
        $enabled = get_option('ugsiw_gs_custom_sheet_name', '0');
        $template = get_option('ugsiw_gs_custom_name_template', 'Orders - {month} {year}');
        ?>
        <div style="margin-bottom: 15px;">
            <label style="display: inline-flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
                <input type="checkbox" name="ugsiw_gs_custom_sheet_name" value="1" <?php checked($enabled, '1'); ?> id="ugsiw_custom_sheet_toggle" style="width: 20px; height: 20px;">
                <span style="font-weight: 500; font-size: 14px;">
                    <span class="dashicons dashicons-edit" style="color: #667eea;"></span>
                    Enable Custom Sheet Names
                </span>
            </label>
        </div>
        
        <div id="ugsiw_custom_name_template" style="<?php echo $enabled !== '1' ? 'display: none;' : ''; ?> margin-top: 15px; padding: 20px; background: #f8f9ff; border-radius: 6px; border: 1px solid #e0e0e0;">
            <input type="text" name="ugsiw_gs_custom_name_template" value="<?php echo esc_attr($template); ?>" style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px;" placeholder="Enter custom sheet name template">
            
            <div style="margin-top: 15px; padding: 15px; background: #fff; border-radius: 4px; border: 1px solid #e0e0e0;">
                <h4 style="margin-top: 0; margin-bottom: 10px; font-size: 14px;">Available Variables:</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; font-size: 12px;">
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{month}</code>
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{year}</code>
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{day}</code>
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{week}</code>
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{site_name}</code>
                    <code style="padding: 5px; background: #f1f1f1; border-radius: 3px;">{order_count}</code>
                </div>
                <p style="margin-top: 10px; margin-bottom: 0; font-size: 12px; color: #666;">Example: <code>{site_name} - {month} {year}</code> would create "My Store - January 2024"</p>
            </div>
        </div>
        
        <p class="description" style="margin-top: 10px;">
            <span class="dashicons dashicons-star-filled" style="color: #ffc107;"></span>
            Pro Feature: Customize sheet names using templates and variables.
        </p>
        
        <script>
        jQuery(document).ready(function($) {
            $('#ugsiw_custom_sheet_toggle').change(function() {
                if ($(this).is(':checked')) {
                    $('#ugsiw_custom_name_template').slideDown();
                } else {
                    $('#ugsiw_custom_name_template').slideUp();
                }
            });
        });
        </script>
        <?php
    }
    
    /**
     * Sanitize array inputs
     */
    public function wpmethods_sanitize_array($input) {
        if (!is_array($input)) {
            return array();
        }
        return array_map('sanitize_text_field', $input);
    }
    
    /**
     * Sanitize checkbox input
     */
    public function wpmethods_sanitize_checkbox($input) {
        return $input ? '1' : '0';
    }
    
    /**
     * Section callback
     */
    public function wpmethods_section_callback() {
        echo '<p>Configure the settings for Google Sheets integration.</p>';
    }

    /**
     * Unified Sheet Mode render (None / Monthly / Daily / Weekly / Product / Custom)
     */
    public function wpmethods_sheet_mode_render() {
        $value = get_option('ugsiw_gs_sheet_mode', 'none');
        $custom_enabled = ($value === 'custom');
        $template = get_option('ugsiw_gs_custom_name_template', 'Orders - {month} {year}');
        // If Pro not active, show locked panel and prompt to upgrade
        if (!$this->is_pro_active) {
            echo '<div style="padding:18px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius:8px; border-left:4px solid #ffc107;">';
            echo '<h4 style="margin:0 0 8px 0;">Sheet Mode <span style="background:#ffc107;color:#663c00;padding:4px 8px;border-radius:16px;font-size:12px;margin-left:8px;">PRO</span></h4>';
            echo '<p style="margin:0 0 8px 0;color:#6b4f00;">Monthly, Daily/Weekly, Product-wise and Custom sheet naming are Pro features. <a href="admin.php?page=ugsiw-license" style="font-weight:700;color:#6b4f00;">Upgrade to Pro</a> to enable them.</p>';
            echo '<div style="display:flex;gap:8px;flex-wrap:wrap;">';
            $modes = array('none' => 'None', 'monthly' => 'Monthly', 'daily' => 'Daily', 'weekly' => 'Weekly', 'product' => 'Product-wise', 'custom' => 'Custom');
            foreach ($modes as $k => $label) {
                $active = ($value === $k) ? 'font-weight:700;color:#222;' : 'color:#7a6a2b;';
                echo '<div style="padding:10px 12px;border-radius:6px;background:#fff;opacity:0.9;border:1px solid #f0e6c8;'.$active.'">'.$label.' <span style="background:#ffc107;color:#663c00;padding:2px 6px;border-radius:12px;font-size:11px;margin-left:6px;">PRO</span></div>';
            }
            echo '</div>';
            echo '</div>';
            return;
        }
        ?>
        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="none" <?php checked($value, 'none'); ?>>
                <span style="font-weight: 500;">None</span>
            </label>

            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="monthly" <?php checked($value, 'monthly'); ?>>
                <span style="font-weight: 500;">Monthly Sheets</span>
            </label>

            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="daily" <?php checked($value, 'daily'); ?>>
                <span style="font-weight: 500;">Daily Sheets</span>
            </label>

            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="weekly" <?php checked($value, 'weekly'); ?>>
                <span style="font-weight: 500;">Weekly Sheets</span>
            </label>

            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="product" <?php checked($value, 'product'); ?>>
                <span style="font-weight: 500;">Product-wise Sheets</span>
            </label>

            <label style="display: flex; align-items: center; gap: 10px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0; cursor: pointer;">
                <input type="radio" name="ugsiw_gs_sheet_mode" value="custom" <?php checked($value, 'custom'); ?> id="ugsiw_sheet_mode_custom">
                <span style="font-weight: 500;">Custom Sheet Naming</span>
            </label>
        </div>

        <div id="ugsiw_sheet_mode_custom_template" style="<?php echo $custom_enabled ? '' : 'display:none;'; ?> margin-top: 15px; padding: 15px; background: #f8f9ff; border-radius: 6px; border: 1px solid #e0e0e0;">
            <input type="text" name="ugsiw_gs_custom_name_template" value="<?php echo esc_attr($template); ?>" style="width:100%; padding:10px; border:1px solid #ddd; border-radius:4px;">
            <p class="description" style="margin-top:8px;">Available variables: {month}, {year}, {day}, {week}, {site_name}, {order_count}</p>
        </div>

        <script>
        jQuery(document).ready(function($){
            $('input[name="ugsiw_gs_sheet_mode"]').change(function(){
                if ($(this).val() === 'custom') {
                    $('#ugsiw_sheet_mode_custom_template').slideDown();
                } else {
                    $('#ugsiw_sheet_mode_custom_template').slideUp();
                }
            });
        });
        </script>
        <?php
    }
    
    /**
     * Monthly sheets field render
     */
    public function wpmethods_monthly_sheets_render() {
        $value = get_option('ugsiw_gs_monthly_sheets', '0');
        ?>
        <label style="display: inline-flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
            <input type="checkbox" name="ugsiw_gs_monthly_sheets" value="1" <?php checked($value, '1'); ?> style="width: 20px; height: 20px;">
            <span style="font-weight: 500; font-size: 14px;">
                <span class="dashicons dashicons-calendar-alt" style="color: #667eea;"></span>
                Enable automatic monthly sheet creation
            </span>
        </label>
        <p class="description" style="margin-top: 10px;">When enabled, a new sheet will be automatically created for each month (e.g., "January 2024", "February 2024")</p>
        <?php
    }

    /**
     * Forward to webhook render (Pro)
     */
    public function wpmethods_forward_webhook_render() {
        if (!$this->is_pro_active) {
            echo '<div style="padding: 20px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">';
            echo '<p style="margin: 0; color: #856404; font-weight: 600;">';
            echo '<span class="dashicons dashicons-lock" style="color: #ffc107;"></span> ';
            echo 'This is a Pro feature. <a href="admin.php?page=ugsiw-license" style="color: #856404; text-decoration: underline;">Upgrade to Pro</a> to enable webhook forwarding.';
            echo '</p>';
            echo '</div>';
            return;
        }

        $enabled = get_option('ugsiw_gs_forward_webhook', '0');
        $url = get_option('ugsiw_gs_webhook_url', '');
        ?>
        <div style="display:flex;gap:12px;align-items:center;">
            <label style="display:inline-flex;align-items:center;gap:8px;padding:10px;background:#f8f9fa;border-radius:6px;border:1px solid #e0e0e0;">
                <input type="checkbox" name="ugsiw_gs_forward_webhook" value="1" <?php checked($enabled, '1'); ?> style="width:18px;height:18px;">
                <span style="font-weight:500;">Enable forwarding to webhook</span>
            </label>
            <input type="url" name="ugsiw_gs_webhook_url" value="<?php echo esc_attr($url); ?>" placeholder="https://hooks.example.com/xyz" style="flex:1;padding:10px;border:1px solid #ddd;border-radius:6px;">
        </div>
        <p class="description" style="margin-top:8px;">When enabled, order data will also be POSTed as JSON to the webhook URL you provide. Useful for Zapier, Integromat, or custom endpoints.</p>
        <?php
    }
    
    /**
     * Order Statuses field render
     */
    public function wpmethods_order_statuses_render() {
        $selected_statuses = get_option('ugsiw_gs_order_statuses', array('completed', 'processing'));
        
        if (!is_array($selected_statuses)) {
            $selected_statuses = maybe_unserialize($selected_statuses);
            if (!is_array($selected_statuses)) {
                $selected_statuses = array('completed', 'processing');
            }
        }
        
        $all_statuses = $this->wpmethods_get_wc_order_statuses();
        
        echo '<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;">';
        
        foreach ($all_statuses as $status => $label) {
            $checked = in_array($status, $selected_statuses) ? 'checked' : '';
            ?>
            <div style="padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
                <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                    <input type="checkbox" name="ugsiw_gs_order_statuses[]" 
                           value="<?php echo esc_attr($status); ?>" <?php echo esc_attr($checked); ?> style="width: 18px; height: 18px;">
                    <span style="font-weight: 500; font-size: 14px;"><?php echo esc_html($label); ?></span>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description" style="margin-top: 10px;">Select order statuses that should trigger sending data to Google Sheets</p>';
    }
    
    /**
     * Get all WooCommerce product categories
     */
    private function wpmethods_get_product_categories() {
        $categories = get_terms(array(
            'taxonomy' => 'product_cat',
            'hide_empty' => false,
            'orderby' => 'name',
            'order' => 'ASC',
        ));
        
        $category_list = array();
        
        if (!is_wp_error($categories) && !empty($categories)) {
            foreach ($categories as $category) {
                $category_list[$category->term_id] = $category->name;
            }
        }
        
        return $category_list;
    }
    
    /**
     * Product Categories field render
     */
    public function wpmethods_product_categories_render() {
        $selected_categories = $this->wpmethods_get_selected_categories();
        $all_categories = $this->wpmethods_get_product_categories();
        
        if (empty($all_categories)) {
            echo '<p>No product categories found.</p>';
            return;
        }
        
        echo '<div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; background: #f8f9fa; border-radius: 6px;">';
        
        foreach ($all_categories as $cat_id => $cat_name) {
            $checked = in_array($cat_id, $selected_categories) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 8px; padding: 10px; background: white; border-radius: 4px; border: 1px solid #e0e0e0;">
                <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                    <input type="checkbox" name="ugsiw_gs_product_categories[]" 
                           value="<?php echo esc_attr($cat_id); ?>" <?php echo esc_attr($checked); ?> style="width: 18px; height: 18px;">
                    <span style="font-weight: 500; font-size: 14px;"><?php echo esc_html($cat_name); ?></span>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description" style="margin-top: 10px;">Select product categories. Orders will only be sent if they contain at least one product from selected categories. Leave empty to include all categories.</p>';
    }
    
    /**
     * Selected fields render
     */
    public function wpmethods_selected_fields_render() {
        $selected_fields = $this->wpmethods_get_selected_fields();
        // Section header with PRO legend
        echo '<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">';
        echo '<h2 style="margin:0;">Checkout Fields</h2>';
        echo '<span style="background:#667eea;color:#fff;padding:4px 8px;border-radius:12px;font-size:12px;">Fields</span>';
        if (!$this->is_pro_active) {
            echo '<span style="margin-left:8px;background:#ffc107;color:#663c00;padding:4px 8px;border-radius:12px;font-size:12px;">PRO fields are locked</span>';
        } else {
            echo '<span style="margin-left:8px;background:#28a745;color:#fff;padding:4px 8px;border-radius:12px;font-size:12px;">PRO active</span>';
        }
        echo '</div>';
        
        // Search box
        ?>
        <div class="wpmethods-search-box">
            <input type="text" id="wpmethods-field-search" placeholder="Search fields...">
        </div>
        <?php
        
        // Display all fields in categories
        echo '<div class="wpmethods-field-category">';
        echo '<h3>';
        echo '<span><span class="dashicons dashicons-list-view"></span> All Fields</span>';
        echo '<span class="dashicons dashicons-arrow-down"></span>';
        echo '</h3>';
        echo '<div class="wpmethods-fields-grid">';
        
        foreach ($this->available_fields as $field_key => $field_info) {
            $checked = in_array($field_key, $selected_fields) ? 'checked' : '';
            $required = (isset($field_info['required']) && $field_info['required']) ? ' <span class="required">Required</span>' : '';
            $is_pro_field = isset($field_info['pro']) && $field_info['pro'];
            $disabled = '';
            $pro_badge = '';

            // Disable Required fields
            if (isset($field_info['required']) && $field_info['required']) {
                $disabled = 'disabled';
            }
            
            if ($is_pro_field && !$this->is_pro_active) {
                $disabled = 'disabled';
                $pro_badge = ' <span style="background:#ffc107;color:#663c00;padding:3px 6px;border-radius:12px;font-size:11px;margin-left:8px;">PRO</span>';
            }
            $icon = isset($field_info['icon']) ? $field_info['icon'] : 'dashicons dashicons-admin-generic';
            ?>
            <div class="wpmethods-field-item">
                <label>
                    <input type="checkbox" name="ugsiw_gs_selected_fields[]" 
                           value="<?php echo esc_attr($field_key); ?>" 
                           <?php echo esc_attr($checked); ?> <?php echo esc_attr($disabled); ?>>
                    <span class="<?php echo esc_attr($icon); ?>" style="color: #667eea;"></span>
                    <span><?php echo esc_html($field_info['label']); ?></span>
                    <?php echo $required; ?><?php echo $pro_badge; ?>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '</div>';
        
        echo '<p class="description" style="margin-top: 20px;">Select fields to include in Google Sheets. Required fields are always included and cannot be disabled.</p>';
        ?>
        <div class="wpmethods-feature-box">
            <h3 style="margin-top: 0; color: #667eea;">
                <span class="dashicons dashicons-code-standards"></span> Generate Google Apps Script
            </h3>
            <p>Click the button below to generate a Google Apps Script code based on your selected fields. This script will handle data submission to your Google Sheets.</p>
            <button type="button" id="wpmethods-generate-script" class="wpmethods-button">
                <span class="dashicons dashicons-update"></span> Generate Google Apps Script
            </button>
            <div id="wpmethods-script-output" style="margin-top: 15px; display: none;">
                <h4>Generated Script</h4>
                <textarea id="wpmethods-generated-script" style="width: 100%; height: 400px;" readonly></textarea>
                <p style="margin-top: 10px;">
                    <button type="button" id="wpmethods-copy-script" class="wpmethods-button wpmethods-button-secondary">
                        <span class="dashicons dashicons-clipboard"></span> Copy to Clipboard
                    </button>
                    <span id="wpmethods-copy-status" style="margin-left: 10px; color: #28a745; font-weight: 600; display: none;">
                        <span class="dashicons dashicons-yes-alt"></span> Copied!
                    </span>
                </p>
            </div>
        </div>
        <?php
    }
    
    /**
     * Script URL field render
     */
    public function wpmethods_script_url_render() {
        $value = get_option('ugsiw_gs_script_url', '');
        ?>
        <div style="max-width: 600px;">
            <input type="url" name="ugsiw_gs_script_url" 
                   value="<?php echo esc_url($value); ?>" 
                   style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px;" 
                   placeholder="https://script.google.com/macros/s/...">
            <p class="description" style="margin-top: 10px;">
                <span class="dashicons dashicons-info" style="color: #667eea;"></span>
                Enter your Google Apps Script web app URL. Get this from Google Apps Script deployment after copying the generated script.
            </p>
        </div>
        <?php
    }
    
    /**
     * AJAX handler for generating Google Apps Script
     */
    public function wpmethods_generate_google_script_ajax() {
        check_ajax_referer('wpmethods_generate_script_nonce', 'nonce');
        
        if (!current_user_can('manage_options')) {
            wp_die('Unauthorized');
        }
        
        $allowed_fields  = array_keys($this->available_fields);
        $selected_fields = $this->wpmethods_get_selected_fields();

        if ( isset($_POST['fields']) && is_array($_POST['fields']) ) {
            $selected_fields = array_intersect(
                array_map('sanitize_key', wp_unslash($_POST['fields'])),
                $allowed_fields
            );
        }

        // Determine sheet mode (single/monthly/daily/weekly/product/custom)
        $sheet_mode = get_option('ugsiw_gs_sheet_mode', 'none');
        $monthly_sheets = ($sheet_mode === 'monthly') ? '1' : '0';
        $daily_weekly = ($sheet_mode === 'daily') ? 'daily' : (($sheet_mode === 'weekly') ? 'weekly' : 'none');
        $product_sheets = ($sheet_mode === 'product') ? '1' : '0';
        $custom_sheet_name = ($sheet_mode === 'custom') ? '1' : '0';
        $custom_template = get_option('ugsiw_gs_custom_name_template', 'Orders - {month} {year}');

        $pro_features = array(
            'daily_weekly'    => $daily_weekly,
            'product_sheets'  => ($product_sheets === '1' || $product_sheets === 1 || $product_sheets === true),
            'custom_sheet_name'=> ($custom_sheet_name === '1' || $custom_sheet_name === 1 || $custom_sheet_name === true),
            'custom_template' => $custom_template,
        );

        // If Pro is not active, force single-sheet mode and remove any Pro fields from selection
        if (!$this->is_pro_active) {
            $sheet_mode = 'none';
            $monthly_sheets = '0';
            $daily_weekly = 'none';
            $product_sheets = '0';
            $custom_sheet_name = '0';
            $pro_features = array(
                'daily_weekly' => 'none',
                'product_sheets' => false,
                'custom_sheet_name' => false,
                'custom_template' => '',
            );

            // Remove pro-only fields from selected fields so generator and payload won't include them
            foreach ($this->available_fields as $fkey => $finfo) {
                if (isset($finfo['pro']) && $finfo['pro']) {
                    if (($k = array_search($fkey, $selected_fields)) !== false) {
                        unset($selected_fields[$k]);
                    }
                }
            }
            // Reindex array
            $selected_fields = array_values($selected_fields);
        }

        // Generate Google Apps Script code using the separate generator (pass pro features)
        $script = $this->script_generator->generate_script($selected_fields, ($monthly_sheets === '1'), $pro_features);

        wp_send_json_success(array(
            'script' => $script,
            'fields' => $selected_fields,
            'monthly_sheets' => $monthly_sheets,
            'pro_features' => $pro_features,
            'sheet_mode' => $sheet_mode,
        ));
    }
    
    
    /**
     * Settings page with modern design
     */
     public function wpmethods_settings_page() {
        $selected_statuses = get_option('ugsiw_gs_order_statuses', array());
        $selected_fields = $this->wpmethods_get_selected_fields();
        $selected_categories = $this->wpmethods_get_selected_categories();

        // Sheet mode (single/monthly/daily/weekly/product/custom)
        $sheet_mode = get_option('ugsiw_gs_sheet_mode', 'none');
        $monthly_sheets = ($sheet_mode === 'monthly') ? '1' : '0';
        // Pro features derived from sheet mode
        $daily_weekly = ($sheet_mode === 'daily') ? 'daily' : (($sheet_mode === 'weekly') ? 'weekly' : 'none');
        $product_sheets = ($sheet_mode === 'product') ? '1' : '0';
        $custom_sheet_name = ($sheet_mode === 'custom') ? '1' : '0';
        ?>
        <div class="wrap wpmethods-settings-wrapper">
            
            <!-- Modern Header with Pro badge -->
            <div class="wpmethods-header">
                <h1>
                    <span class="dashicons dashicons-google" style="vertical-align: middle; margin-right: 10px;"></span>
                    WooCommerce to Google Sheets
                    <?php if ($this->is_pro_active): ?>
                    <span style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-size: 12px; padding: 4px 10px; border-radius: 20px; vertical-align: middle; margin-left: 10px;">PRO</span>
                    <?php endif; ?>
                </h1>
                <p>Automatically send WooCommerce orders to Google Sheets. Configure your integration below.</p>
                
                <?php if (!$this->is_pro_active): ?>
                <div style="background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); padding: 15px; border-radius: 8px; margin-top: 15px; border-left: 4px solid #ffc107; max-width: 600px;">
                    <p style="margin: 0; color: #856404; font-weight: 600;">
                        <span class="dashicons dashicons-unlock" style="color: #ffc107;"></span>
                        Want Daily/Weekly Sheets, Product-wise Sheets, and Custom Sheet Naming? 
                        <a href="admin.php?page=ugsiw-license" style="color: #856404; text-decoration: underline; font-weight: 700;">Upgrade to Pro</a>
                    </p>
                </div>
                <?php endif; ?>
            </div>
            
            <!-- Dashboard Stats -->
            <div class="wpmethods-dashboard">
                <div class="wpmethods-card">
                    <h3><span class="dashicons dashicons-admin-settings"></span> Configuration Status</h3>
                    <div class="wpmethods-stats-grid">
                        <div class="wpmethods-stat">
                            <div class="wpmethods-stat-number"><?php echo count($selected_fields); ?></div>
                            <div class="wpmethods-stat-label">Selected Fields</div>
                        </div>
                        <div class="wpmethods-stat">
                            <div class="wpmethods-stat-number"><?php echo count($selected_statuses); ?></div>
                            <div class="wpmethods-stat-label">Trigger Statuses</div>
                        </div>
                        <div class="wpmethods-stat">
                            <?php
                                $mode_label = 'Single';
                                $checked = '✓';
                                if ($sheet_mode === 'monthly') {
                                    $mode_label = 'Monthly Sheets';
                                    $checked = '✓';
                                } elseif ($sheet_mode === 'daily') {
                                    $mode_label = 'Daily Sheets';
                                    $checked = '✓';
                                } elseif ($sheet_mode === 'weekly') {
                                    $mode_label = 'Weekly Sheets';
                                    $checked = '✓';
                                } elseif ($sheet_mode === 'product') {
                                    $mode_label = 'Product-wise Sheets';
                                    $checked = '✓';
                                } elseif ($sheet_mode === 'custom') {
                                    $mode_label = 'Custom Sheet Naming';
                                    $checked = '✓';
                                }
                            ?>
                            <div class="wpmethods-stat-number"><?php echo $checked; ?></div>
                            <div class="wpmethods-stat-label"><?php echo esc_html($mode_label); ?></div>
                        </div>
                        
                    </div>
                </div>
                
                <!-- Pro Features Card -->
                <?php if (!$this->is_pro_active): ?>
                <div class="wpmethods-card" style="border: 2px solid #ffc107;">
                    <h3 style="color: #ffc107;">
                        <span class="dashicons dashicons-star-filled"></span> Unlock Pro Features
                    </h3>
                    <ul style="margin: 0; padding-left: 20px; color: #666;">
                        <li style="margin-bottom: 8px;">✅ Daily/Weekly Sheets - Auto creation</li>
                        <li style="margin-bottom: 8px;">✅ Product-wise Sheets - Category based</li>
                        <li style="margin-bottom: 8px;">✅ Custom Sheet Naming - Flexible templates</li>
                        <li style="margin-bottom: 8px;">✅ Advanced Field Mapping - Extra fields</li>
                        <li style="margin-bottom: 8px;">✅ Priority Support - Fast help</li>
                    </ul>
                    <a href="admin.php?page=ugsiw-license" class="wpmethods-button" style="margin-top: 15px; background: linear-gradient(135deg, #ffc107 0%, #ff9800 100%);">
                        <span class="dashicons dashicons-unlock"></span> Upgrade to Pro
                    </a>
                </div>
                <?php endif; ?>
            </div>
            
            <!-- Main Settings Form -->
            <form action="options.php" method="post" style="margin-bottom: 30px;">
                <?php
                settings_fields('ugsiw_gs_settings');
                do_settings_sections('ugsiw_gs_settings');
                submit_button('Save Settings', 'primary wpmethods-button', 'submit', false);
                ?>
            </form>
            
            
            
            <!-- Configuration Summary -->
            <div class="wpmethods-card">
                <h3><span class="dashicons dashicons-chart-bar"></span> Configuration Summary</h3>
                
                <div class="wpmethods-stats-grid">
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number">
                            <?php 
                            if (!is_array($selected_statuses)) {
                                $selected_statuses = maybe_unserialize($selected_statuses);
                                if (!is_array($selected_statuses)) {
                                    $selected_statuses = array('completed', 'processing');
                                }
                            }
                            echo count($selected_statuses);
                            ?>
                        </div>
                        <div class="wpmethods-stat-label">Trigger Statuses</div>
                    </div>
                    
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number"><?php echo $monthly_sheets === '1' ? 'Enabled' : 'Single'; ?></div>
                        <div class="wpmethods-stat-label">Sheet Mode</div>
                    </div>
                    
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number"><?php echo count($selected_fields); ?></div>
                        <div class="wpmethods-stat-label">Selected Fields</div>
                    </div>
                </div>
                
                <div style="margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 6px;">
                    <h4 style="margin-top: 0; margin-bottom: 10px;">Selected Fields:</h4>
                    <p style="margin: 0; color: #666;">
                        <?php 
                        $field_names = array();
                        foreach ($selected_fields as $field_key) {
                            if (isset($this->available_fields[$field_key])) {
                                $field_names[] = $this->available_fields[$field_key]['label'];
                            }
                        }
                        echo esc_html(implode(', ', $field_names));
                        ?>
                    </p>
                </div>

                <?php if ($monthly_sheets === '1'): ?>
                    <div style="margin-top: 20px; padding: 15px; background: #e8f5e9; border-radius: 6px; border-left: 4px solid #4CAF50;">
                        <h4 style="margin-top: 0; color: #2e7d32;">
                            <span class="dashicons dashicons-calendar-alt"></span> Monthly Sheets Active
                        </h4>
                        <p style="margin: 10px 0 0 0; color: #2e7d32;">
                            Orders will be automatically organized into monthly sheets
                            (
                            <?php
                                echo esc_html( gmdate('F Y') );
                                echo ', ';
                                echo esc_html( gmdate('F Y', strtotime('+1 month', time())) );
                            ?>,
                            etc.)
                        </p>
                    </div>
                <?php endif; ?>
            </div>


            <!-- Video Tutorial -->
            <div class="wpmethods-video-box">
                <h3><span class="dashicons dashicons-video-alt3"></span> Watch Setup Tutorial</h3>
                <p style="color: rgba(255,255,255,0.9); margin-bottom: 20px;">Learn how to set up the plugin step by step with our video tutorial.</p>
                <a href="https://youtu.be/7Kh-uugbods" target="_blank" class="wpmethods-video-button">
                    <span class="dashicons dashicons-youtube"></span> Watch Tutorial Video
                </a>
            </div>


            <!-- More Products Section -->
            <div class="wpmethods-products-box">
                <h3><span class="dashicons dashicons-products"></span> More Products by WP Methods</h3>
                <p style="margin-bottom: 20px;">Check out our other products that might interest you.</p>
                
                <div class="wpmethods-products-grid">
                    <!-- Product 1 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2025/01/Click-Shop-wordpress-theme.jpg" alt="Click Shop - Ecommerce Wordpress Theme">
                        <h4>Click Shop - Ecommerce Wordpress Theme</h4>
                        <a href="https://wpmethods.com/product/click-shop-wordpress-landing-page-type-ecommerce-theme/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>
                    
                    <!-- Product 2 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2025/07/bdexchanger-dollar-buy-sell-php-script-or-money-exchanger-ex.webp" alt="BDExchanger || PHP Script for Dollar Buy Sell">
                        <h4>BDExchanger || PHP Script for Dollar Buy Sell</h4>
                        <a href="https://wpmethods.com/product/bdexchanger-php-script-for-dollar-buy-sell-or-currency-exchanger/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>
                    
                    <!-- Product 3 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2025/05/Social-Chat-Floating-Icons-WordPress-Plugin.webp" alt="Social Chat Floating Icons Wordpress Plugin">
                        <h4>Social Chat Floating Icons Wordpress Plugin</h4>
                        <a href="https://wpmethods.com/product/social-chat-floating-icons-wordpress-plugin/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>
                    
                    <!-- Product 4 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2022/12/How-to-Show-Recent-WooCommerce-Order-List-Table-with-Elementor-Addon-Orders-Frontend.jpg" alt="WooCommerce Order List Table on eCommerce Website">
                        <h4>WooCommerce Order List Table on eCommerce Website</h4>
                        <a href="https://wpmethods.com/product/woocommerce-order-list-table-on-ecommerce-website-elementor-addon/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>
                    
                    <!-- Product 5 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2024/12/Book-Shop-Multi-Seller-banner.jpg" alt="Multi-Vendor Book Selling Website Backup File">
                        <h4>Multi-Vendor Book Selling Website Backup File</h4>
                        <a href="https://wpmethods.com/product/multi-vendor-book-selling-website-to-sell-pdf-hardcover-books/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>

                    <!-- Product 6 -->
                    <div class="wpmethods-product-card">
                        <img src="https://wpmethods.com/media/2023/10/single-product-landing-page-with-woocommerce-checkout-form-copy.jpg" alt="Multi-Vendor Book Selling Website Backup File">
                        <h4>Single Product Landing Page with WooCommerce</h4>
                        <a href="https://wpmethods.com/product/single-product-landing-page-with-woocommerce-checkout-order-form/" target="_blank" class="wpmethods-product-button">Get it</a>
                    </div>
                </div>
                
                <div style="text-align: center; margin-top: 30px;">
                    <a href="https://wpmethods.com/" target="_blank" class="wpmethods-button">View More Products</a>
                </div>
            </div>


            
            <!-- Donation Section -->
            <div class="wpmethods-donation-box">
                <h3><span class="dashicons dashicons-heart"></span> Support This Plugin</h3>
                <p style="color: #856404; margin-bottom: 20px;">If this plugin has helped your business, consider buying me a coffee to support further development.</p>
                <a href="https://buymeacoffee.com/ajharrashed" target="_blank" class="wpmethods-donation-button">
                    <span class="dashicons dashicons-coffee"></span> Buy Me a Coffee
                </a>
            </div>
            
        </div>
        <?php
    }
}