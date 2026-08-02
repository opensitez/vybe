<?php
// vybe-test: php/array_advanced/array_reduce_complex
// origin: languages/php/tests/php/test_array_advanced.rs

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

$items = [
    ["name" => "apple", "price" => 1.5, "qty" => 3],
    ["name" => "banana", "price" => 0.75, "qty" => 6],
    ["name" => "cherry", "price" => 2.0, "qty" => 2],
];
$total = array_reduce($items, function($carry, $item) {
    return $carry + ($item["price"] * $item["qty"]);
}, 0);
echo $total;

__vybe_check(ob_get_clean(), "13");
