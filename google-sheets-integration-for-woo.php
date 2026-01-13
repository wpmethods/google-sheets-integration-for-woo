<?php
/**
 * Plugin Name: Google Sheets Integration for WooCommerce by WP Methods
 * Plugin URI: https://wpmethods.com/plugins/google-sheets-integration-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 2.0.5
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 * License: GPL2
 * Text Domain: wpmethods-wc-to-gs
 * Domain Path: /languages
 * Requires at least: 5.9
 * Requires PHP: 7.4
 */

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

class WPMethods_WC_To_Google_Sheets {
    
    private $available_fields = array();
    
    public function __construct() {
        // Define available fields
        $this->define_available_fields();
        
        // Check if WooCommerce is active
        add_action('admin_init', array($this, 'wpmethods_check_woocommerce'));
        
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
                'always_include' => true
            ),
            'billing_name' => array(
                'label' => 'Billing Name',
                'required' => true,
                'always_include' => true
            ),
            'billing_email' => array(
                'label' => 'Email Address',
                'required' => false,
                'always_include' => false
            ),
            'billing_phone' => array(
                'label' => 'Phone',
                'required' => false,
                'always_include' => false
            ),
            'billing_address' => array(
                'label' => 'Billing Address',
                'required' => false,
                'always_include' => false
            ),
            'product_name' => array(
                'label' => 'Product Name',
                'required' => true,
                'always_include' => true
            ),
            'order_amount_with_currency' => array(
                'label' => 'Order Amount',
                'required' => true,
                'always_include' => true
            ),
            'order_currency' => array(
                'label' => 'Order Currency',
                'required' => false,
                'always_include' => false
            ),
            'order_status' => array(
                'label' => 'Order Status',
                'required' => true,
                'always_include' => true
            ),
            'order_date' => array(
                'label' => 'Order Date',
                'required' => true,
                'always_include' => true
            ),
            'product_categories' => array(
                'label' => 'Product Categories',
                'required' => false,
                'always_include' => false
            )
        );
    }
    
    /**
     * Get selected fields for Google Sheets
     */
    private function wpmethods_get_selected_fields() {
        $selected_fields = get_option('wpmethods_wc_gs_selected_fields', array());
        
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
            <p><?php _e('WP Methods WooCommerce to Google Sheets requires WooCommerce to be installed and activated.', 'wpmethods-wc-to-gs'); ?></p>
        </div>
        <?php
    }
    
    /**
     * Plugin activation - set default values
     */
    public function wpmethods_activate_plugin() {
        if (!get_option('wpmethods_wc_gs_order_statuses')) {
            update_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        }
        
        if (!get_option('wpmethods_wc_gs_script_url')) {
            update_option('wpmethods_wc_gs_script_url', '');
        }
        
        if (!get_option('wpmethods_wc_gs_product_categories')) {
            update_option('wpmethods_wc_gs_product_categories', array());
        }
        
        // Set default selected fields (all fields)
        if (!get_option('wpmethods_wc_gs_selected_fields')) {
            $default_fields = array_keys($this->available_fields);
            update_option('wpmethods_wc_gs_selected_fields', $default_fields);
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
        $selected_categories = get_option('wpmethods_wc_gs_product_categories', array());
        
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
     * Check if order contains products from selected categories
     */
    private function wpmethods_order_has_selected_categories($order, $selected_categories) {
        if (empty($selected_categories)) {
            return true;
        }
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                $product_categories = wp_get_post_terms($product->get_id(), 'product_cat', array('fields' => 'ids'));
                
                if (!empty(array_intersect($product_categories, $selected_categories))) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * Send order data to Google Sheets
     */
    public function wpmethods_send_order_to_sheets($order_id, $old_status, $new_status, $order) {
        
        $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        
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
        
        $script_url = get_option('wpmethods_wc_gs_script_url', '');
        
        if (empty($script_url)) {
            error_log('WP Methods Google Sheets: Google Apps Script URL not configured');
            return;
        }
        
        $order_data = $this->wpmethods_prepare_order_data($order);
        
        $response = wp_remote_post($script_url, array(
            'method' => 'POST',
            'timeout' => 45,
            'redirection' => 5,
            'httpversion' => '1.0',
            'blocking' => true,
            'headers' => array(
                'Content-Type' => 'application/json',
            ),
            'body' => json_encode($order_data),
            'cookies' => array()
        ));
        
        if (is_wp_error($response)) {
            error_log('WP Methods Google Sheets integration error: ' . $response->get_error_message());
        } else {
            // Debug logging for successful submissions
            error_log('WP Methods Google Sheets: Order ' . $order_id . ' sent successfully');
        }
    }
    
    /**
     * Prepare order data for Google Sheets based on selected fields
     */
    private function wpmethods_prepare_order_data($order) {
        $selected_fields = $this->wpmethods_get_selected_fields();
        $order_data = array();
        
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
                
            case 'order_status':
                return $order->get_status();
                
            case 'order_date':
                $date_created = $order->get_date_created();
                return $date_created ? $date_created->format('Y-m-d H:i:s') : '';
                
            case 'product_categories':
                return $this->wpmethods_get_order_categories($order);
                
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
                $categories = wp_get_post_terms($product->get_id(), 'product_cat', array('fields' => 'names'));
                $all_categories = array_merge($all_categories, $categories);
            }
        }
        
        $all_categories = array_unique($all_categories);
        
        return implode(', ', $all_categories);
    }
    
    /**
     * Add admin menu
     */
    public function wpmethods_add_admin_menu() {
        add_options_page(
            'WooCommerce to Google Sheets',
            'WC to Google Sheets',
            'manage_options',
            'wpmethods-wc-to-google-sheets',
            array($this, 'wpmethods_settings_page')
        );
    }
    
    /**
     * Admin scripts
     */
    public function wpmethods_admin_scripts($hook) {
        if ($hook != 'settings_page_wpmethods-wc-to-google-sheets') {
            return;
        }
        
        // Create admin.js content inline
        $admin_js = '
        jQuery(document).ready(function($) {
            // Generate Google Apps Script
            $("#wpmethods-generate-script").on("click", function(e) {
                e.preventDefault();
                
                var button = $(this);
                var originalText = button.text();
                
                button.text("Generating...").prop("disabled", true);
                
                // Get selected fields
                var selectedFields = [];
                $("input[name=\'wpmethods_wc_gs_selected_fields[]\']:checked").each(function() {
                    selectedFields.push($(this).val());
                });
                
                $.ajax({
                    url: ajaxurl,
                    type: "POST",
                    data: {
                        action: "wpmethods_generate_google_script",
                        nonce: "' . wp_create_nonce('wpmethods_generate_script_nonce') . '",
                        fields: selectedFields
                    },
                    success: function(response) {
                        if (response.success) {
                            $("#wpmethods-generated-script").val(response.data.script);
                            $("#wpmethods-script-output").show();
                            $("html, body").animate({
                                scrollTop: $("#wpmethods-script-output").offset().top - 100
                            }, 500);
                        } else {
                            alert("Error generating script. Please try again.");
                        }
                    },
                    error: function() {
                        alert("Error generating script. Please try again.");
                    },
                    complete: function() {
                        button.text(originalText).prop("disabled", false);
                    }
                });
            });
            
            // Copy to clipboard
            $("#wpmethods-copy-script").on("click", function() {
                var textarea = $("#wpmethods-generated-script")[0];
                textarea.select();
                document.execCommand("copy");
                
                $("#wpmethods-copy-status").show().fadeOut(2000);
            });
            
            // Handle required fields
            $("input[name=\'wpmethods_wc_gs_selected_fields[]\']").each(function() {
                if ($(this).is(":disabled")) {
                    $(this).prop("checked", true);
                }
            });
            
            // Prevent unchecking required fields
            $("input[name=\'wpmethods_wc_gs_selected_fields[]\']").on("change", function() {
                if ($(this).is(":disabled") && !$(this).is(":checked")) {
                    $(this).prop("checked", true);
                }
            });
        });
        ';
        
        // Add inline script
        wp_add_inline_script('jquery', $admin_js);
        
        // Add inline CSS
        $admin_css = '
        .wpmethods-settings-wrapper {
            max-width: 1200px;
        }
        .wpmethods-field-group {
            background: #f9f9f9;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .wpmethods-field-group h3 {
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 1px solid #ddd;
        }
        .wpmethods-field-checkboxes {
            max-height: 300px;
            overflow-y: auto;
            padding: 10px;
        }
        .wpmethods-field-item {
            margin-bottom: 8px;
            padding: 5px;
            background: white;
            border-radius: 3px;
        }
        .wpmethods-field-item label {
            display: flex;
            align-items: center;
            cursor: pointer;
        }
        .wpmethods-field-item input[type="checkbox"] {
            margin-right: 8px;
        }
        .wpmethods-field-item .required {
            color: #d63638;
            font-weight: bold;
            margin-left: 5px;
        }
        .wpmethods-generated-script {
            font-family: "Courier New", monospace;
            font-size: 12px;
            line-height: 1.4;
        }
        #wpmethods-script-output {
            animation: fadeIn 0.3s ease-in;
        }
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        ';
        
        wp_add_inline_style('wp-admin', $admin_css);
    }
    
    /**
     * Initialize settings
     */
    public function wpmethods_settings_init() {
        // Main settings
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_order_statuses', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_script_url', 'esc_url_raw');
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_product_categories', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_selected_fields', array($this, 'wpmethods_sanitize_array'));
        
        // Main settings section
        add_settings_section(
            'wpmethods_wc_gs_section',
            'Google Sheets Integration Settings',
            array($this, 'wpmethods_section_callback'),
            'wpmethods_wc_gs_settings'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_order_statuses',
            'Trigger Order Statuses',
            array($this, 'wpmethods_order_statuses_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
        
        
        
        add_settings_field(
            'wpmethods_wc_gs_product_categories',
            'Product Categories Filter',
            array($this, 'wpmethods_product_categories_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );

        add_settings_field(
            'wpmethods_wc_gs_selected_fields',
            'Checkout Fields',
            array($this, 'wpmethods_selected_fields_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_script_url',
            'Google Apps Script URL',
            array($this, 'wpmethods_script_url_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
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
     * Section callback
     */
    public function wpmethods_section_callback() {
        echo '<p>Configure the settings for Google Sheets integration.</p>';
    }
    
    /**
     * Order Statuses field render
     */
    public function wpmethods_order_statuses_render() {
        $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        
        if (!is_array($selected_statuses)) {
            $selected_statuses = maybe_unserialize($selected_statuses);
            if (!is_array($selected_statuses)) {
                $selected_statuses = array('completed', 'processing');
            }
        }
        
        $all_statuses = $this->wpmethods_get_wc_order_statuses();
        
        foreach ($all_statuses as $status => $label) {
            $checked = in_array($status, $selected_statuses) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 5px;">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_order_statuses[]" 
                           value="<?php echo esc_attr($status); ?>" <?php echo $checked; ?>>
                    <?php echo esc_html($label); ?>
                </label>
            </div>
            <?php
        }
        echo '<p class="description">Select order statuses that should trigger sending data to Google Sheets</p>';
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
        
        echo '<div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;">';
        
        foreach ($all_categories as $cat_id => $cat_name) {
            $checked = in_array($cat_id, $selected_categories) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 5px;">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_product_categories[]" 
                           value="<?php echo esc_attr($cat_id); ?>" <?php echo $checked; ?>>
                    <?php echo esc_html($cat_name); ?>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description">Select product categories. Orders will only be sent if they contain at least one product from selected categories. Leave empty to include all categories.</p>';
    }
    
    /**
     * Selected fields render
     */
    public function wpmethods_selected_fields_render() {
        $selected_fields = $this->wpmethods_get_selected_fields();
        
        echo '<h4>Checkout Fields</h4>';
        echo '<div style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin-bottom: 20px;">';
        
        foreach ($this->available_fields as $field_key => $field_info) {
            $checked = in_array($field_key, $selected_fields) ? 'checked' : '';
            $disabled = (isset($field_info['always_include']) && $field_info['always_include']) ? 'disabled' : '';
            $required = (isset($field_info['required']) && $field_info['required']) ? ' <span style="color:red">*</span>' : '';
            ?>
            <div style="margin-bottom: 5px;">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_selected_fields[]" 
                           value="<?php echo esc_attr($field_key); ?>" 
                           <?php echo $checked; ?> <?php echo $disabled; ?>>
                    <?php echo esc_html($field_info['label']) . $required; ?>
                    <?php if ($disabled): ?>
                        <em>(Required)</em>
                    <?php endif; ?>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description">Select fields to include in Google Sheets. Required fields (*) are always included.</p>';
        ?>
        <div style="margin-top: 20px; padding: 15px; background: #f5f5f5; border-radius: 4px;">
            <h4>Generate Google Apps Script</h4>
            <p>Click the button below to generate a Google Apps Script code based on your selected fields:</p>
            <button type="button" id="wpmethods-generate-script" class="button button-primary">
                Generate Google Apps Script
            </button>
            <div id="wpmethods-script-output" style="margin-top: 15px; display: none;">
                <textarea id="wpmethods-generated-script" style="width: 100%; height: 400px; font-family: monospace;" readonly></textarea>
                <p style="margin-top: 10px;">
                    <button type="button" id="wpmethods-copy-script" class="button button-secondary">
                        Copy to Clipboard
                    </button>
                    <span id="wpmethods-copy-status" style="margin-left: 10px; color: green; display: none;">Copied!</span>
                </p>
            </div>
        </div>
        <?php
    }
    
    /**
     * Script URL field render
     */
    public function wpmethods_script_url_render() {
        $value = get_option('wpmethods_wc_gs_script_url', '');
        ?>
        <input type="url" name="wpmethods_wc_gs_script_url" 
               value="<?php echo esc_url($value); ?>" 
               style="width: 500px;" 
               placeholder="https://script.google.com/macros/s/...">
        <p class="description">Enter your Google Apps Script web app URL. Get this from Google Apps Script deployment.</p>
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
        
        $selected_fields = isset($_POST['fields']) ? (array) $_POST['fields'] : $this->wpmethods_get_selected_fields();
        
        // Generate Google Apps Script code
        $script = $this->generate_google_apps_script($selected_fields);
        
        wp_send_json_success(array(
            'script' => $script,
            'fields' => $selected_fields
        ));
    }
    
    /**
     * Generate Google Apps Script code based on selected fields
     */
    private function generate_google_apps_script($selected_fields) {
        // Get field labels for headers
        $headers = array();
        $field_mapping = array();
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $headers[] = $this->available_fields[$field_key]['label'];
                $field_mapping[$field_key] = $this->available_fields[$field_key]['label'];
            }
        }
        
        $headers_js = json_encode($headers, JSON_PRETTY_PRINT);
        $field_mapping_js = json_encode($field_mapping, JSON_PRETTY_PRINT);
        
        $script = <<<EOT
// Google Apps Script Code for Google Sheets
// Generated by WP Methods WooCommerce to Google Sheets Plugin
// Fields: {$this->get_field_list($selected_fields)}

function doPost(e) {
    try {
        // Parse the incoming data
        const data = JSON.parse(e.postData.contents);
        
        // Get the active sheet
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        
        // Initialize headers if sheet is empty
        initializeSheet(sheet);
        
        // Check if this order already exists
        const orderIds = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
        const existingRowIndex = orderIds.indexOf(data.order_id.toString());
        
        if (existingRowIndex !== -1) {
            // Update existing row
            updateExistingRow(sheet, existingRowIndex, data);
        } else {
            // Add new row
            addNewRow(sheet, data);
        }
        
        // Return success response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Order data saved successfully'
        })).setMimeType(ContentService.MimeType.JSON);
        
    } catch (error) {
        // Return error response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function initializeSheet(sheet) {
    if (sheet.getLastRow() === 0) {
        const headers = {$headers_js};
        sheet.appendRow(headers);
    }
}

function updateExistingRow(sheet, existingRowIndex, data) {
    const row = existingRowIndex + 2; // +2 for header row and 0-based index
    
    const fieldOrder = {$headers_js};
    
    fieldOrder.forEach((fieldLabel, index) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        if (data[fieldKey] !== undefined) {
            sheet.getRange(row, index + 1).setValue(data[fieldKey]);
        }
    });
}

function addNewRow(sheet, data) {
    const fieldOrder = {$headers_js};
    const rowData = [];
    
    fieldOrder.forEach((fieldLabel) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        rowData.push(data[fieldKey] || '');
    });
    
    sheet.appendRow(rowData);
}

function getFieldKeyFromLabel(fieldLabel) {
    const fieldMap = {$field_mapping_js};
    
    // Reverse lookup: find key by label
    for (const [key, label] of Object.entries(fieldMap)) {
        if (label === fieldLabel) {
            return key;
        }
    }
    
    // Fallback: convert label to lowercase with underscores
    return fieldLabel.toLowerCase().replace(/ /g, '_');
}

// Function to manually initialize the sheet with headers
function manualInitialize() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    initializeSheet(sheet);
}
EOT;

        return $script;
    }
    
    /**
     * Get field list string
     */
    private function get_field_list($selected_fields) {
        $field_names = array();
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $field_names[] = $this->available_fields[$field_key]['label'];
            }
        }
        return implode(', ', $field_names);
    }
    
    /**
     * Settings page
     */
    public function wpmethods_settings_page() {
        ?>
        <div class="wrap">
            <h1>WooCommerce to Google Sheets</h1>
            
            <form action="options.php" method="post">
                <?php
                settings_fields('wpmethods_wc_gs_settings');
                do_settings_sections('wpmethods_wc_gs_settings');
                submit_button();
                ?>
            </form>
            
            <h3>Current Configuration Summary:</h3>
            <?php
            $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array());
            
            if (!is_array($selected_statuses)) {
                $selected_statuses = maybe_unserialize($selected_statuses);
                if (!is_array($selected_statuses)) {
                    $selected_statuses = array('completed', 'processing');
                }
            }
            
            echo '<p><strong>Trigger Statuses:</strong> ';
            if (!empty($selected_statuses)) {
                $status_labels = array();
                foreach ($selected_statuses as $status) {
                    $status_labels[] = ucfirst($status);
                }
                echo implode(', ', $status_labels);
            } else {
                echo 'None selected';
            }
            echo '</p>';
            
            $selected_fields = $this->wpmethods_get_selected_fields();
            echo '<p><strong>Selected Fields (' . count($selected_fields) . '):</strong> ';
            $field_names = array();
            foreach ($selected_fields as $field_key) {
                if (isset($this->available_fields[$field_key])) {
                    $field_names[] = $this->available_fields[$field_key]['label'];
                }
            }
            echo implode(', ', $field_names);
            echo '</p>';
            
            $selected_categories = $this->wpmethods_get_selected_categories();
            $all_categories = $this->wpmethods_get_product_categories();
            
            echo '<p><strong>Selected Categories:</strong> ';
            if (!empty($selected_categories)) {
                $category_names = array();
                foreach ($selected_categories as $cat_id) {
                    if (isset($all_categories[$cat_id])) {
                        $category_names[] = $all_categories[$cat_id];
                    }
                }
                echo !empty($category_names) ? implode(', ', $category_names) : 'All categories';
            } else {
                echo 'All categories (no filter)';
            }
            echo '</p>';
            ?>
        </div>
        <?php
    }
}

// Initialize the plugin
new WPMethods_WC_To_Google_Sheets();