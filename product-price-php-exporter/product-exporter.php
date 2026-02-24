<?php
require_once 'wp-load.php';

header('Content-Type: text/csv');
header('Content-Disposition: attachment;filename=products.csv');

$output = fopen('php://output', 'w');
fputcsv($output, ['Name', 'Sale Price']);

$args = [
    'limit' => -1,
    'status' => 'publish',
];

$products = wc_get_products($args);

foreach ($products as $product) {

    if ($product->is_type('variable')) {
        $variations = $product->get_children();
        if (!empty($variations)) {
            $variation = wc_get_product($variations[0]); // فقط اولین variation
            fputcsv($output, [
                $variation->get_name(),
                $variation->get_sale_price()
            ]);
        }
    } else {
        fputcsv($output, [
            $product->get_name(),
            $product->get_sale_price()
        ]);
    }
}

fclose($output);
