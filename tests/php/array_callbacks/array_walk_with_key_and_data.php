<?php
// vybe-test: php/array_callbacks/array_walk_with_key_and_data
// origin: languages/php/tests/php/test_array_callbacks.rs

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

$items = ['x' => 1, 'y' => 2];
$factor = 5;
array_walk($items, function(&$v, $k, $factor) {
    $v *= $factor;
    if ($k === 'x') {
        $v += 1;
    }
}, $factor);
echo implode('|', $items);

__vybe_check(ob_get_clean(), "6|10");
