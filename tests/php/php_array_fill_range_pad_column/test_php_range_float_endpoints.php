<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_range_float_endpoints
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs

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

$r = range(0.5, 2.5, 1);
echo implode("|", $r);

__vybe_check(ob_get_clean(), "0.5|1.5|2.5");
