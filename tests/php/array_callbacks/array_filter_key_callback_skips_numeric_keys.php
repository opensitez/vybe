<?php
// vybe-test: php/array_callbacks/array_filter_key_callback_skips_numeric_keys
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

$items = ['0' => 1, 'a' => 2, 3 => 3, 'b' => 4];
$filtered = array_filter($items, fn($k) => ctype_alpha((string)$k), ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($filtered));

__vybe_check(ob_get_clean(), "a,b");
