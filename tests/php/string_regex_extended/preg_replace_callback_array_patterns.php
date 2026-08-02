<?php
// vybe-test: php/string_regex_extended/preg_replace_callback_array_patterns
// origin: languages/php/tests/php/test_string_regex_extended.rs

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

$r = preg_replace_callback_array([
    '/\b[A-Z][a-z]+/' => fn($m) => strtolower($m[0]),
    '/\b\d+/' => fn($m) => $m[0] * 10,
], 'Hello 5 World 3');
echo $r;
echo "\n";

__vybe_check(ob_get_clean(), "hello 50 world 30");
