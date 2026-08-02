<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_returns_matching_key
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs

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

$map = ["first" => 10, "target" => 20, "third" => 30];
if (function_exists('array_find_key')) {
    $key = array_find_key($map, fn($val) => $val === 20);
    echo "Found Key: $key";
} else {
    $key = null;
    foreach ($map as $k => $v) {
        if ($v === 20) { $key = $k; break; }
    }
    echo "Found Key: $key";
}

__vybe_check(ob_get_clean(), "Found Key: target");
