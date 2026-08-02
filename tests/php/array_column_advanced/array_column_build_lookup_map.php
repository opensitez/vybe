<?php
// vybe-test: php/array_column_advanced/array_column_build_lookup_map
// origin: languages/php/tests/php/test_array_column_advanced.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$products = [
    ['sku'=>'A001','price'=>9.99],
    ['sku'=>'B002','price'=>14.99],
    ['sku'=>'C003','price'=>4.99],
];
$prices = array_column($products, 'price', 'sku');
echo $prices['B002'];

__vybe_check(ob_get_clean(), "14.99");
