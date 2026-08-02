<?php
// vybe-test: php/regex_patterns/char_class_patterns
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

echo preg_match('/^\d+$/', "12345") ? "digits" : "no";
echo preg_match('/^[a-zA-Z]+$/', "Hello") ? "alpha" : "no";
echo preg_match('/^[\w]+$/', "hello_123") ? "word" : "no";
echo preg_match('/^[^aeiou]+$/i', "fly") ? "no vowels" : "has vowels";

__vybe_check(ob_get_clean(), "digitsalphawordno vowels");
