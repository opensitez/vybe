<?php
// vybe-test: php/regex_patterns/preg_replace_callback_basic
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

$result = preg_replace_callback('/\d+/', function($matches) {
    return $matches[0] * 2;
}, "I have 5 apples and 3 oranges");
echo $result;

__vybe_check(ob_get_clean(), "I have 10 apples and 6 oranges");
