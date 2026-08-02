<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_mb_convert_case_multibyte_modes_runtime
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs

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

echo mb_convert_case("straße", MB_CASE_UPPER, "UTF-8");
echo "|";
echo mb_convert_case("İ", MB_CASE_LOWER, "UTF-8");

__vybe_check(ob_get_clean(), "STRASSE|i̇");
