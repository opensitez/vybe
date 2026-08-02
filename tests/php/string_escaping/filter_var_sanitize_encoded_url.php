<?php
// vybe-test: php/string_escaping/filter_var_sanitize_encoded_url
// origin: languages/php/tests/php/test_string_escaping.rs

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

$out = filter_var('https://example.com/q?a=1', FILTER_SANITIZE_URL);
echo $out;

__vybe_check(ob_get_clean(), "https://example.com/q?a=1");
