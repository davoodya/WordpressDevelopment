<?php
/*
Plugin Name: Hello Davood
Description: افزونه آموزشی اولیه — نمایش شورتکد و یک گزینه در منو.
Version: 0.1
Author: Davood
Text Domain: hello-davood
*/

# TODO: 1. Add Settings 
if (! defined('ABSPATH')) exit;

// ثبت شورتکد ساده
function hd_hello_shortcode($atts)
{
    $atts = shortcode_atts(array(
        'name' => 'دوست',
    ), $atts);
    return sprintf('<div class="hd-hello">سلام %s — از افزونه Hello Davood!</div>', esc_html($atts['name']));
}
add_shortcode('hello_davood', 'hd_hello_shortcode');

// یک صفحه تنظیمات ساده در منوی ادمین
function hd_add_admin_menu()
{
    add_menu_page('Hello Davood', 'HelloDav', 'manage_options', 'hd-main', 'hd_settings_page', 'dashicons-smiley', 60);
}
add_action('admin_menu', 'hd_add_admin_menu');

function hd_settings_page()
{
?>
    <div class="wrap">
        <h1><?php esc_html_e('Hello Davood Settings', 'hello-davood'); ?></h1>
        <p><?php esc_html_e('این یک نمونه صفحه تنظیمات است.', 'hello-davood'); ?></p>
    </div>
<?php
}
