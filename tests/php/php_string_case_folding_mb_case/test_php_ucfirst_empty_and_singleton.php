<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_ucfirst_empty_and_singleton
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

echo ucfirst("") === "" ? "empty" : "no";
echo "|";
echo ucfirst("δ");
echo "|";
echo ucfirst("ß");

__vybe_check(ob_get_clean(), "empty|δ|ß");
