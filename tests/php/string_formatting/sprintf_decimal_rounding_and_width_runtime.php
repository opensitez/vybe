<?php
// vybe-test: php/string_formatting/sprintf_decimal_rounding_and_width_runtime
// origin: languages/php/tests/php/test_string_formatting.rs

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

echo sprintf("%.1f", 1.24);
echo "|";
echo sprintf("%.1f", 1.25);
echo "|";
echo sprintf("%+08d", -42);
echo "|";
echo sprintf("%#.3x", 255);

__vybe_check(ob_get_clean(), "1.2|1.3|-0000042|0xfff");
