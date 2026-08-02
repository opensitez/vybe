<?php
// vybe-test: php/regex_patterns/validate_email_pattern
// origin: languages/php/tests/php/test_regex_patterns.rs

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

$pattern = '/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/';
echo preg_match($pattern, "user@example.com") ? "valid" : "invalid";
echo preg_match($pattern, "bad@") ? "valid" : "invalid";
echo preg_match($pattern, "test.user+tag@domain.co.uk") ? "valid" : "invalid";

__vybe_check(ob_get_clean(), "validinvalidvalid");
