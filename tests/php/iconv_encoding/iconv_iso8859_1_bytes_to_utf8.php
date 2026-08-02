<?php
// vybe-test: php/iconv_encoding/iconv_iso8859_1_bytes_to_utf8
// origin: languages/php/tests/php/test_iconv_encoding.rs

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

echo iconv('ISO-8859-1', 'UTF-8', "\xE9");

__vybe_check(ob_get_clean(), "é");
