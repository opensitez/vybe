<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_strstr_needle_not_found_runtime
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs

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

echo mb_strstr("abcdef", "z") === false ? "missing" : "found";
echo "|";
echo mb_strrchr("caféc", "é") === "éc" ? "tail" : "no";

__vybe_check(ob_get_clean(), "missing|tail");
