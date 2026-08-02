<?php
// vybe-test: php/array_map_multiple/array_filter_mode_both
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

$r = array_filter(['x' => 1, 'y' => 2], fn($v,$k) => $k === 'x' || $v > 1, ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($r));

__vybe_check(ob_get_clean(), "x,y");
