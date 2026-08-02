<?php
// vybe-test: php/numeric_casts/bcdiv_by_zero_returns_null_without_scale
// origin: languages/php/tests/php/test_numeric_casts.rs

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

$v = bcdiv('1', '0');
echo $v === null ? 'null' : (string)$v;

__vybe_check(ob_get_clean(), "null");
