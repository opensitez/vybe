<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_replace_callback_callable
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs

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

$input = "word1 word2 word3";
$result = preg_replace_callback('/\b\w+\b/', function($matches) {
    return strtoupper($matches[0]);
}, $input);
echo $result;

__vybe_check(ob_get_clean(), "WORD1 WORD2 WORD3");
