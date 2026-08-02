<?php
// vybe-test: php/array_column_object_properties/array_column_objects_public_map_with_indexing
// origin: languages/php/tests/php/test_array_column_object_properties.rs

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

class Product {
    public function __construct(public string $sku, public float $price) {}
}
$rows = [
    new Product('X', 10.5),
    new Product('Y', 12.0),
];
$prices = array_column($rows, 'price', 'sku');
echo $prices['X'] . "|" . $prices['Y'];

__vybe_check(ob_get_clean(), "10.5|12");
