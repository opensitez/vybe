<?php
// vybe-test: php/string_advanced/str_getcsv_limits_and_empty_fields_runtime
// origin: languages/php/tests/php/test_string_advanced.rs

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

$v = str_getcsv("a, c,");
echo count($v);
echo "|";
echo $v[1];
echo "|";
echo $v[3];
echo "\n";
$u = str_getcsv('"a","b","c"', ',', '"', '\\');
echo $u[2];
echo "|";
echo implode('-', $u);

__vybe_check(ob_get_clean(), "4||\nc|a-b-c");
