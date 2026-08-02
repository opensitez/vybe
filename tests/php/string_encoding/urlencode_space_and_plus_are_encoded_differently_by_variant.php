<?php
// vybe-test: php/string_encoding/urlencode_space_and_plus_are_encoded_differently_by_variant
// origin: languages/php/tests/php/test_string_encoding.rs

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

$raw = 'a b+c';
echo urlencode($raw);
echo '|';
echo rawurlencode($raw);

__vybe_check(ob_get_clean(), "a+b%2Bc|a%20b%2Bc");
