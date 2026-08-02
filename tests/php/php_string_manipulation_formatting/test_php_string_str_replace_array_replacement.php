<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_str_replace_array_replacement
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs

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

$vowels = ["a", "e", "i", "o", "u"];
$res = str_replace($vowels, "*", "Hello World");
echo $res;

__vybe_check(ob_get_clean(), "H*ll* W*rld");
