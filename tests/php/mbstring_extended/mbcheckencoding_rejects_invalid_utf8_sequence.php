<?php
// vybe-test: php/mbstring_extended/mbcheckencoding_rejects_invalid_utf8_sequence
// origin: languages/php/tests/php/test_mbstring_extended.rs

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

echo mb_check_encoding("\xFF\xFE", 'UTF-8') ? 'valid' : 'invalid';

__vybe_check(ob_get_clean(), "invalid");
