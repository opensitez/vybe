<?php
// vybe-test: php/array_callbacks/array_filter_with_empty_callback_and_both_modes
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

$items = ['x' => 0, 'y' => 1, 'z' => 2];
$filtered = array_filter(
    $items,
    fn($v, $k) => $v > 0 || $k === 'x',
    ARRAY_FILTER_USE_BOTH
);
echo implode(',', array_keys($filtered));

__vybe_check(ob_get_clean(), "x,y,z");
